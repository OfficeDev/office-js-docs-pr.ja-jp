---
title: Office. メールボックス-要件セット1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 10ee3c6f83488c777ae5e06dbb9484382a9c9588
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268657"
---
# <a name="item"></a><span data-ttu-id="97316-102">item</span><span class="sxs-lookup"><span data-stu-id="97316-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="97316-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="97316-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="97316-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="97316-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-106">要件</span><span class="sxs-lookup"><span data-stu-id="97316-106">Requirements</span></span>

|<span data-ttu-id="97316-107">要件</span><span class="sxs-lookup"><span data-stu-id="97316-107">Requirement</span></span>| <span data-ttu-id="97316-108">値</span><span class="sxs-lookup"><span data-stu-id="97316-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-110">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-110">1.0</span></span>|
|[<span data-ttu-id="97316-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="97316-112">Restricted</span></span>|
|[<span data-ttu-id="97316-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="97316-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="97316-115">Members and methods</span></span>

| <span data-ttu-id="97316-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-116">Member</span></span> | <span data-ttu-id="97316-117">種類</span><span class="sxs-lookup"><span data-stu-id="97316-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="97316-118">attachments</span><span class="sxs-lookup"><span data-stu-id="97316-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="97316-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-119">Member</span></span> |
| [<span data-ttu-id="97316-120">bcc</span><span class="sxs-lookup"><span data-stu-id="97316-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="97316-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-121">Member</span></span> |
| [<span data-ttu-id="97316-122">body</span><span class="sxs-lookup"><span data-stu-id="97316-122">body</span></span>](#body-body) | <span data-ttu-id="97316-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-123">Member</span></span> |
| [<span data-ttu-id="97316-124">cc</span><span class="sxs-lookup"><span data-stu-id="97316-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="97316-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-125">Member</span></span> |
| [<span data-ttu-id="97316-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="97316-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="97316-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-127">Member</span></span> |
| [<span data-ttu-id="97316-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="97316-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="97316-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-129">Member</span></span> |
| [<span data-ttu-id="97316-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="97316-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="97316-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-131">Member</span></span> |
| [<span data-ttu-id="97316-132">end</span><span class="sxs-lookup"><span data-stu-id="97316-132">end</span></span>](#end-datetime) | <span data-ttu-id="97316-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-133">Member</span></span> |
| [<span data-ttu-id="97316-134">from</span><span class="sxs-lookup"><span data-stu-id="97316-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="97316-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-135">Member</span></span> |
| [<span data-ttu-id="97316-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="97316-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="97316-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-137">Member</span></span> |
| [<span data-ttu-id="97316-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="97316-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="97316-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-139">Member</span></span> |
| [<span data-ttu-id="97316-140">itemId</span><span class="sxs-lookup"><span data-stu-id="97316-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="97316-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-141">Member</span></span> |
| [<span data-ttu-id="97316-142">itemType</span><span class="sxs-lookup"><span data-stu-id="97316-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="97316-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-143">Member</span></span> |
| [<span data-ttu-id="97316-144">location</span><span class="sxs-lookup"><span data-stu-id="97316-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="97316-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-145">Member</span></span> |
| [<span data-ttu-id="97316-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="97316-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="97316-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-147">Member</span></span> |
| [<span data-ttu-id="97316-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="97316-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="97316-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-149">Member</span></span> |
| [<span data-ttu-id="97316-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="97316-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="97316-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-151">Member</span></span> |
| [<span data-ttu-id="97316-152">organizer</span><span class="sxs-lookup"><span data-stu-id="97316-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="97316-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-153">Member</span></span> |
| [<span data-ttu-id="97316-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="97316-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="97316-155">Member</span><span class="sxs-lookup"><span data-stu-id="97316-155">Member</span></span> |
| [<span data-ttu-id="97316-156">sender</span><span class="sxs-lookup"><span data-stu-id="97316-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="97316-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-157">Member</span></span> |
| [<span data-ttu-id="97316-158">start</span><span class="sxs-lookup"><span data-stu-id="97316-158">start</span></span>](#start-datetime) | <span data-ttu-id="97316-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-159">Member</span></span> |
| [<span data-ttu-id="97316-160">subject</span><span class="sxs-lookup"><span data-stu-id="97316-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="97316-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-161">Member</span></span> |
| [<span data-ttu-id="97316-162">to</span><span class="sxs-lookup"><span data-stu-id="97316-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="97316-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-163">Member</span></span> |
| [<span data-ttu-id="97316-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97316-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="97316-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-165">Method</span></span> |
| [<span data-ttu-id="97316-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97316-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="97316-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-167">Method</span></span> |
| [<span data-ttu-id="97316-168">close</span><span class="sxs-lookup"><span data-stu-id="97316-168">close</span></span>](#close) | <span data-ttu-id="97316-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-169">Method</span></span> |
| [<span data-ttu-id="97316-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="97316-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="97316-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-171">Method</span></span> |
| [<span data-ttu-id="97316-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="97316-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="97316-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-173">Method</span></span> |
| [<span data-ttu-id="97316-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="97316-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="97316-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-175">Method</span></span> |
| [<span data-ttu-id="97316-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="97316-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="97316-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-177">Method</span></span> |
| [<span data-ttu-id="97316-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="97316-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="97316-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-179">Method</span></span> |
| [<span data-ttu-id="97316-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="97316-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="97316-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-181">Method</span></span> |
| [<span data-ttu-id="97316-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="97316-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="97316-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-183">Method</span></span> |
| [<span data-ttu-id="97316-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="97316-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="97316-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-185">Method</span></span> |
| [<span data-ttu-id="97316-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="97316-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="97316-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-187">Method</span></span> |
| [<span data-ttu-id="97316-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="97316-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="97316-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-189">Method</span></span> |
| [<span data-ttu-id="97316-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="97316-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="97316-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-191">Method</span></span> |
| [<span data-ttu-id="97316-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="97316-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="97316-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="97316-194">例</span><span class="sxs-lookup"><span data-stu-id="97316-194">Example</span></span>

<span data-ttu-id="97316-195">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="97316-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="97316-196">メンバー</span><span class="sxs-lookup"><span data-stu-id="97316-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="97316-197">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="97316-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="97316-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-200">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="97316-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="97316-201">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97316-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="97316-202">型</span><span class="sxs-lookup"><span data-stu-id="97316-202">Type</span></span>

*   <span data-ttu-id="97316-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="97316-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-204">要件</span><span class="sxs-lookup"><span data-stu-id="97316-204">Requirements</span></span>

|<span data-ttu-id="97316-205">要件</span><span class="sxs-lookup"><span data-stu-id="97316-205">Requirement</span></span>| <span data-ttu-id="97316-206">値</span><span class="sxs-lookup"><span data-stu-id="97316-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-208">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-208">1.0</span></span>|
|[<span data-ttu-id="97316-209">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-210">ReadItem</span></span>|
|[<span data-ttu-id="97316-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-212">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-213">例</span><span class="sxs-lookup"><span data-stu-id="97316-213">Example</span></span>

<span data-ttu-id="97316-214">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="97316-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="97316-215">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-216">メッセージの Bcc (ブラインドカーボンコピー) 行を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="97316-217">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-218">型</span><span class="sxs-lookup"><span data-stu-id="97316-218">Type</span></span>

*   [<span data-ttu-id="97316-219">受信者</span><span class="sxs-lookup"><span data-stu-id="97316-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-220">要件</span><span class="sxs-lookup"><span data-stu-id="97316-220">Requirements</span></span>

|<span data-ttu-id="97316-221">要件</span><span class="sxs-lookup"><span data-stu-id="97316-221">Requirement</span></span>| <span data-ttu-id="97316-222">値</span><span class="sxs-lookup"><span data-stu-id="97316-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-224">1.1</span><span class="sxs-lookup"><span data-stu-id="97316-224">1.1</span></span>|
|[<span data-ttu-id="97316-225">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-226">ReadItem</span></span>|
|[<span data-ttu-id="97316-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-228">作成</span><span class="sxs-lookup"><span data-stu-id="97316-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-229">例</span><span class="sxs-lookup"><span data-stu-id="97316-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="97316-230">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-231">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-232">型</span><span class="sxs-lookup"><span data-stu-id="97316-232">Type</span></span>

*   [<span data-ttu-id="97316-233">Body</span><span class="sxs-lookup"><span data-stu-id="97316-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-234">要件</span><span class="sxs-lookup"><span data-stu-id="97316-234">Requirements</span></span>

|<span data-ttu-id="97316-235">要件</span><span class="sxs-lookup"><span data-stu-id="97316-235">Requirement</span></span>| <span data-ttu-id="97316-236">値</span><span class="sxs-lookup"><span data-stu-id="97316-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-238">1.1</span><span class="sxs-lookup"><span data-stu-id="97316-238">1.1</span></span>|
|[<span data-ttu-id="97316-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-240">ReadItem</span></span>|
|[<span data-ttu-id="97316-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-242">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-243">例</span><span class="sxs-lookup"><span data-stu-id="97316-243">Example</span></span>

<span data-ttu-id="97316-244">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="97316-245">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="97316-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="97316-246">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-247">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="97316-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="97316-248">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-249">Read mode</span></span>

<span data-ttu-id="97316-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-252">Compose mode</span></span>

<span data-ttu-id="97316-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97316-254">型</span><span class="sxs-lookup"><span data-stu-id="97316-254">Type</span></span>

*   <span data-ttu-id="97316-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-256">要件</span><span class="sxs-lookup"><span data-stu-id="97316-256">Requirements</span></span>

|<span data-ttu-id="97316-257">要件</span><span class="sxs-lookup"><span data-stu-id="97316-257">Requirement</span></span>| <span data-ttu-id="97316-258">値</span><span class="sxs-lookup"><span data-stu-id="97316-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-260">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-260">1.0</span></span>|
|[<span data-ttu-id="97316-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-262">ReadItem</span></span>|
|[<span data-ttu-id="97316-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-264">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="97316-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="97316-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="97316-266">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="97316-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="97316-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="97316-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-271">型</span><span class="sxs-lookup"><span data-stu-id="97316-271">Type</span></span>

*   <span data-ttu-id="97316-272">String</span><span class="sxs-lookup"><span data-stu-id="97316-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-273">要件</span><span class="sxs-lookup"><span data-stu-id="97316-273">Requirements</span></span>

|<span data-ttu-id="97316-274">要件</span><span class="sxs-lookup"><span data-stu-id="97316-274">Requirement</span></span>| <span data-ttu-id="97316-275">値</span><span class="sxs-lookup"><span data-stu-id="97316-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-277">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-277">1.0</span></span>|
|[<span data-ttu-id="97316-278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-279">ReadItem</span></span>|
|[<span data-ttu-id="97316-280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-281">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-282">例</span><span class="sxs-lookup"><span data-stu-id="97316-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="97316-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="97316-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="97316-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-286">型</span><span class="sxs-lookup"><span data-stu-id="97316-286">Type</span></span>

*   <span data-ttu-id="97316-287">日付</span><span class="sxs-lookup"><span data-stu-id="97316-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-288">要件</span><span class="sxs-lookup"><span data-stu-id="97316-288">Requirements</span></span>

|<span data-ttu-id="97316-289">要件</span><span class="sxs-lookup"><span data-stu-id="97316-289">Requirement</span></span>| <span data-ttu-id="97316-290">値</span><span class="sxs-lookup"><span data-stu-id="97316-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-292">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-292">1.0</span></span>|
|[<span data-ttu-id="97316-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-294">ReadItem</span></span>|
|[<span data-ttu-id="97316-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-297">例</span><span class="sxs-lookup"><span data-stu-id="97316-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="97316-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="97316-298">dateTimeModified: Date</span></span>

<span data-ttu-id="97316-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-301">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-302">型</span><span class="sxs-lookup"><span data-stu-id="97316-302">Type</span></span>

*   <span data-ttu-id="97316-303">日付</span><span class="sxs-lookup"><span data-stu-id="97316-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-304">要件</span><span class="sxs-lookup"><span data-stu-id="97316-304">Requirements</span></span>

|<span data-ttu-id="97316-305">要件</span><span class="sxs-lookup"><span data-stu-id="97316-305">Requirement</span></span>| <span data-ttu-id="97316-306">値</span><span class="sxs-lookup"><span data-stu-id="97316-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-308">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-308">1.0</span></span>|
|[<span data-ttu-id="97316-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-310">ReadItem</span></span>|
|[<span data-ttu-id="97316-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-313">例</span><span class="sxs-lookup"><span data-stu-id="97316-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="97316-314">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="97316-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="97316-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-318">Read mode</span></span>

<span data-ttu-id="97316-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-320">Compose mode</span></span>

<span data-ttu-id="97316-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="97316-322">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="97316-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="97316-323">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="97316-324">型</span><span class="sxs-lookup"><span data-stu-id="97316-324">Type</span></span>

*   <span data-ttu-id="97316-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-326">要件</span><span class="sxs-lookup"><span data-stu-id="97316-326">Requirements</span></span>

|<span data-ttu-id="97316-327">要件</span><span class="sxs-lookup"><span data-stu-id="97316-327">Requirement</span></span>| <span data-ttu-id="97316-328">値</span><span class="sxs-lookup"><span data-stu-id="97316-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-330">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-330">1.0</span></span>|
|[<span data-ttu-id="97316-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-332">ReadItem</span></span>|
|[<span data-ttu-id="97316-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="97316-335">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="97316-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="97316-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-340">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="97316-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-341">型</span><span class="sxs-lookup"><span data-stu-id="97316-341">Type</span></span>

*   [<span data-ttu-id="97316-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97316-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-343">要件</span><span class="sxs-lookup"><span data-stu-id="97316-343">Requirements</span></span>

|<span data-ttu-id="97316-344">要件</span><span class="sxs-lookup"><span data-stu-id="97316-344">Requirement</span></span>| <span data-ttu-id="97316-345">値</span><span class="sxs-lookup"><span data-stu-id="97316-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-347">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-347">1.0</span></span>|
|[<span data-ttu-id="97316-348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-349">ReadItem</span></span>|
|[<span data-ttu-id="97316-350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-351">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-352">例</span><span class="sxs-lookup"><span data-stu-id="97316-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="97316-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="97316-353">internetMessageId: String</span></span>

<span data-ttu-id="97316-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-356">型</span><span class="sxs-lookup"><span data-stu-id="97316-356">Type</span></span>

*   <span data-ttu-id="97316-357">String</span><span class="sxs-lookup"><span data-stu-id="97316-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-358">要件</span><span class="sxs-lookup"><span data-stu-id="97316-358">Requirements</span></span>

|<span data-ttu-id="97316-359">要件</span><span class="sxs-lookup"><span data-stu-id="97316-359">Requirement</span></span>| <span data-ttu-id="97316-360">値</span><span class="sxs-lookup"><span data-stu-id="97316-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-362">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-362">1.0</span></span>|
|[<span data-ttu-id="97316-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-364">ReadItem</span></span>|
|[<span data-ttu-id="97316-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-367">例</span><span class="sxs-lookup"><span data-stu-id="97316-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="97316-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="97316-368">itemClass: String</span></span>

<span data-ttu-id="97316-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="97316-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="97316-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="97316-373">型</span><span class="sxs-lookup"><span data-stu-id="97316-373">Type</span></span> | <span data-ttu-id="97316-374">説明</span><span class="sxs-lookup"><span data-stu-id="97316-374">Description</span></span> | <span data-ttu-id="97316-375">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="97316-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="97316-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="97316-376">Appointment items</span></span> | <span data-ttu-id="97316-377">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="97316-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="97316-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="97316-378">Message items</span></span> | <span data-ttu-id="97316-379">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="97316-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="97316-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="97316-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-381">型</span><span class="sxs-lookup"><span data-stu-id="97316-381">Type</span></span>

*   <span data-ttu-id="97316-382">String</span><span class="sxs-lookup"><span data-stu-id="97316-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-383">要件</span><span class="sxs-lookup"><span data-stu-id="97316-383">Requirements</span></span>

|<span data-ttu-id="97316-384">要件</span><span class="sxs-lookup"><span data-stu-id="97316-384">Requirement</span></span>| <span data-ttu-id="97316-385">値</span><span class="sxs-lookup"><span data-stu-id="97316-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-387">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-387">1.0</span></span>|
|[<span data-ttu-id="97316-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-389">ReadItem</span></span>|
|[<span data-ttu-id="97316-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-392">例</span><span class="sxs-lookup"><span data-stu-id="97316-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="97316-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="97316-393">(nullable) itemId: String</span></span>

<span data-ttu-id="97316-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="97316-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="97316-397">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="97316-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="97316-398">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="97316-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="97316-399">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97316-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="97316-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="97316-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-402">型</span><span class="sxs-lookup"><span data-stu-id="97316-402">Type</span></span>

*   <span data-ttu-id="97316-403">String</span><span class="sxs-lookup"><span data-stu-id="97316-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-404">要件</span><span class="sxs-lookup"><span data-stu-id="97316-404">Requirements</span></span>

|<span data-ttu-id="97316-405">要件</span><span class="sxs-lookup"><span data-stu-id="97316-405">Requirement</span></span>| <span data-ttu-id="97316-406">値</span><span class="sxs-lookup"><span data-stu-id="97316-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-408">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-408">1.0</span></span>|
|[<span data-ttu-id="97316-409">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-410">ReadItem</span></span>|
|[<span data-ttu-id="97316-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-413">例</span><span class="sxs-lookup"><span data-stu-id="97316-413">Example</span></span>

<span data-ttu-id="97316-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="97316-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="97316-416">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-417">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="97316-418">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="97316-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-419">型</span><span class="sxs-lookup"><span data-stu-id="97316-419">Type</span></span>

*   [<span data-ttu-id="97316-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="97316-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-421">要件</span><span class="sxs-lookup"><span data-stu-id="97316-421">Requirements</span></span>

|<span data-ttu-id="97316-422">要件</span><span class="sxs-lookup"><span data-stu-id="97316-422">Requirement</span></span>| <span data-ttu-id="97316-423">値</span><span class="sxs-lookup"><span data-stu-id="97316-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-425">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-425">1.0</span></span>|
|[<span data-ttu-id="97316-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-427">ReadItem</span></span>|
|[<span data-ttu-id="97316-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-429">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-430">例</span><span class="sxs-lookup"><span data-stu-id="97316-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="97316-431">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-433">Read mode</span></span>

<span data-ttu-id="97316-434">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-435">Compose mode</span></span>

<span data-ttu-id="97316-436">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97316-437">型</span><span class="sxs-lookup"><span data-stu-id="97316-437">Type</span></span>

*   <span data-ttu-id="97316-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-439">要件</span><span class="sxs-lookup"><span data-stu-id="97316-439">Requirements</span></span>

|<span data-ttu-id="97316-440">要件</span><span class="sxs-lookup"><span data-stu-id="97316-440">Requirement</span></span>| <span data-ttu-id="97316-441">値</span><span class="sxs-lookup"><span data-stu-id="97316-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-443">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-443">1.0</span></span>|
|[<span data-ttu-id="97316-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-445">ReadItem</span></span>|
|[<span data-ttu-id="97316-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-447">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="97316-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="97316-448">normalizedSubject: String</span></span>

<span data-ttu-id="97316-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="97316-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="97316-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-453">型</span><span class="sxs-lookup"><span data-stu-id="97316-453">Type</span></span>

*   <span data-ttu-id="97316-454">String</span><span class="sxs-lookup"><span data-stu-id="97316-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-455">要件</span><span class="sxs-lookup"><span data-stu-id="97316-455">Requirements</span></span>

|<span data-ttu-id="97316-456">要件</span><span class="sxs-lookup"><span data-stu-id="97316-456">Requirement</span></span>| <span data-ttu-id="97316-457">値</span><span class="sxs-lookup"><span data-stu-id="97316-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-458">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-459">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-459">1.0</span></span>|
|[<span data-ttu-id="97316-460">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-461">ReadItem</span></span>|
|[<span data-ttu-id="97316-462">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-463">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-464">例</span><span class="sxs-lookup"><span data-stu-id="97316-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="97316-465">notificationMessages: [Notificationmessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-466">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-467">型</span><span class="sxs-lookup"><span data-stu-id="97316-467">Type</span></span>

*   [<span data-ttu-id="97316-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="97316-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-469">要件</span><span class="sxs-lookup"><span data-stu-id="97316-469">Requirements</span></span>

|<span data-ttu-id="97316-470">要件</span><span class="sxs-lookup"><span data-stu-id="97316-470">Requirement</span></span>| <span data-ttu-id="97316-471">値</span><span class="sxs-lookup"><span data-stu-id="97316-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-473">1.3</span><span class="sxs-lookup"><span data-stu-id="97316-473">1.3</span></span>|
|[<span data-ttu-id="97316-474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-475">ReadItem</span></span>|
|[<span data-ttu-id="97316-476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-477">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-478">例</span><span class="sxs-lookup"><span data-stu-id="97316-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="97316-479">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-480">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="97316-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="97316-481">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-482">Read mode</span></span>

<span data-ttu-id="97316-483">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-484">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-484">Compose mode</span></span>

<span data-ttu-id="97316-485">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97316-486">型</span><span class="sxs-lookup"><span data-stu-id="97316-486">Type</span></span>

*   <span data-ttu-id="97316-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-488">要件</span><span class="sxs-lookup"><span data-stu-id="97316-488">Requirements</span></span>

|<span data-ttu-id="97316-489">要件</span><span class="sxs-lookup"><span data-stu-id="97316-489">Requirement</span></span>| <span data-ttu-id="97316-490">値</span><span class="sxs-lookup"><span data-stu-id="97316-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-491">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-492">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-492">1.0</span></span>|
|[<span data-ttu-id="97316-493">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-494">ReadItem</span></span>|
|[<span data-ttu-id="97316-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-496">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="97316-497">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-500">型</span><span class="sxs-lookup"><span data-stu-id="97316-500">Type</span></span>

*   [<span data-ttu-id="97316-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97316-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-502">要件</span><span class="sxs-lookup"><span data-stu-id="97316-502">Requirements</span></span>

|<span data-ttu-id="97316-503">要件</span><span class="sxs-lookup"><span data-stu-id="97316-503">Requirement</span></span>| <span data-ttu-id="97316-504">値</span><span class="sxs-lookup"><span data-stu-id="97316-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-506">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-506">1.0</span></span>|
|[<span data-ttu-id="97316-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-508">ReadItem</span></span>|
|[<span data-ttu-id="97316-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-510">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-511">例</span><span class="sxs-lookup"><span data-stu-id="97316-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="97316-512">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-513">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="97316-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="97316-514">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-515">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-515">Read mode</span></span>

<span data-ttu-id="97316-516">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-517">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-517">Compose mode</span></span>

<span data-ttu-id="97316-518">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="97316-519">型</span><span class="sxs-lookup"><span data-stu-id="97316-519">Type</span></span>

*   <span data-ttu-id="97316-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-521">要件</span><span class="sxs-lookup"><span data-stu-id="97316-521">Requirements</span></span>

|<span data-ttu-id="97316-522">要件</span><span class="sxs-lookup"><span data-stu-id="97316-522">Requirement</span></span>| <span data-ttu-id="97316-523">値</span><span class="sxs-lookup"><span data-stu-id="97316-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-525">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-525">1.0</span></span>|
|[<span data-ttu-id="97316-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-527">ReadItem</span></span>|
|[<span data-ttu-id="97316-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="97316-530">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="97316-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="97316-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="97316-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="97316-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="97316-536">型</span><span class="sxs-lookup"><span data-stu-id="97316-536">Type</span></span>

*   [<span data-ttu-id="97316-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="97316-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="97316-538">要件</span><span class="sxs-lookup"><span data-stu-id="97316-538">Requirements</span></span>

|<span data-ttu-id="97316-539">要件</span><span class="sxs-lookup"><span data-stu-id="97316-539">Requirement</span></span>| <span data-ttu-id="97316-540">値</span><span class="sxs-lookup"><span data-stu-id="97316-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-542">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-542">1.0</span></span>|
|[<span data-ttu-id="97316-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-544">ReadItem</span></span>|
|[<span data-ttu-id="97316-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-547">例</span><span class="sxs-lookup"><span data-stu-id="97316-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="97316-548">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="97316-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="97316-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-552">Read mode</span></span>

<span data-ttu-id="97316-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-554">Compose mode</span></span>

<span data-ttu-id="97316-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="97316-556">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="97316-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="97316-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="97316-558">型</span><span class="sxs-lookup"><span data-stu-id="97316-558">Type</span></span>

*   <span data-ttu-id="97316-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-560">要件</span><span class="sxs-lookup"><span data-stu-id="97316-560">Requirements</span></span>

|<span data-ttu-id="97316-561">要件</span><span class="sxs-lookup"><span data-stu-id="97316-561">Requirement</span></span>| <span data-ttu-id="97316-562">値</span><span class="sxs-lookup"><span data-stu-id="97316-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-564">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-564">1.0</span></span>|
|[<span data-ttu-id="97316-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-566">ReadItem</span></span>|
|[<span data-ttu-id="97316-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-568">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="97316-569">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="97316-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="97316-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-572">Read mode</span></span>

<span data-ttu-id="97316-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-575">Compose mode</span></span>

<span data-ttu-id="97316-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="97316-577">型</span><span class="sxs-lookup"><span data-stu-id="97316-577">Type</span></span>

*   <span data-ttu-id="97316-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-579">要件</span><span class="sxs-lookup"><span data-stu-id="97316-579">Requirements</span></span>

|<span data-ttu-id="97316-580">要件</span><span class="sxs-lookup"><span data-stu-id="97316-580">Requirement</span></span>| <span data-ttu-id="97316-581">値</span><span class="sxs-lookup"><span data-stu-id="97316-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-583">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-583">1.0</span></span>|
|[<span data-ttu-id="97316-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-585">ReadItem</span></span>|
|[<span data-ttu-id="97316-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-587">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="97316-588">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="97316-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="97316-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="97316-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="97316-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="97316-591">Read mode</span></span>

<span data-ttu-id="97316-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="97316-594">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="97316-594">Compose mode</span></span>

<span data-ttu-id="97316-595">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="97316-596">型</span><span class="sxs-lookup"><span data-stu-id="97316-596">Type</span></span>

*   <span data-ttu-id="97316-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-598">要件</span><span class="sxs-lookup"><span data-stu-id="97316-598">Requirements</span></span>

|<span data-ttu-id="97316-599">要件</span><span class="sxs-lookup"><span data-stu-id="97316-599">Requirement</span></span>| <span data-ttu-id="97316-600">値</span><span class="sxs-lookup"><span data-stu-id="97316-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-601">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-602">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-602">1.0</span></span>|
|[<span data-ttu-id="97316-603">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-604">ReadItem</span></span>|
|[<span data-ttu-id="97316-605">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-606">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="97316-607">メソッド</span><span class="sxs-lookup"><span data-stu-id="97316-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="97316-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97316-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="97316-609">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="97316-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="97316-610">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="97316-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="97316-611">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="97316-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-612">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-612">Parameters</span></span>

|<span data-ttu-id="97316-613">名前</span><span class="sxs-lookup"><span data-stu-id="97316-613">Name</span></span>| <span data-ttu-id="97316-614">型</span><span class="sxs-lookup"><span data-stu-id="97316-614">Type</span></span>| <span data-ttu-id="97316-615">属性</span><span class="sxs-lookup"><span data-stu-id="97316-615">Attributes</span></span>| <span data-ttu-id="97316-616">説明</span><span class="sxs-lookup"><span data-stu-id="97316-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="97316-617">String</span><span class="sxs-lookup"><span data-stu-id="97316-617">String</span></span>||<span data-ttu-id="97316-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="97316-620">String</span><span class="sxs-lookup"><span data-stu-id="97316-620">String</span></span>||<span data-ttu-id="97316-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="97316-623">Object</span><span class="sxs-lookup"><span data-stu-id="97316-623">Object</span></span>| <span data-ttu-id="97316-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-624">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-626">Object</span><span class="sxs-lookup"><span data-stu-id="97316-626">Object</span></span>| <span data-ttu-id="97316-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-627">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-628">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97316-629">function</span><span class="sxs-lookup"><span data-stu-id="97316-629">function</span></span>| <span data-ttu-id="97316-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-630">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-631">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97316-632">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="97316-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="97316-633">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="97316-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97316-634">エラー</span><span class="sxs-lookup"><span data-stu-id="97316-634">Errors</span></span>

| <span data-ttu-id="97316-635">エラー コード</span><span class="sxs-lookup"><span data-stu-id="97316-635">Error code</span></span> | <span data-ttu-id="97316-636">説明</span><span class="sxs-lookup"><span data-stu-id="97316-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="97316-637">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="97316-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="97316-638">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="97316-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="97316-639">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="97316-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-640">要件</span><span class="sxs-lookup"><span data-stu-id="97316-640">Requirements</span></span>

|<span data-ttu-id="97316-641">要件</span><span class="sxs-lookup"><span data-stu-id="97316-641">Requirement</span></span>| <span data-ttu-id="97316-642">値</span><span class="sxs-lookup"><span data-stu-id="97316-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-643">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-644">1.1</span><span class="sxs-lookup"><span data-stu-id="97316-644">1.1</span></span>|
|[<span data-ttu-id="97316-645">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-647">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-648">作成</span><span class="sxs-lookup"><span data-stu-id="97316-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-649">例</span><span class="sxs-lookup"><span data-stu-id="97316-649">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="97316-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97316-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="97316-651">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="97316-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="97316-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="97316-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="97316-655">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="97316-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="97316-656">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="97316-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-657">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-657">Parameters</span></span>

|<span data-ttu-id="97316-658">名前</span><span class="sxs-lookup"><span data-stu-id="97316-658">Name</span></span>| <span data-ttu-id="97316-659">型</span><span class="sxs-lookup"><span data-stu-id="97316-659">Type</span></span>| <span data-ttu-id="97316-660">属性</span><span class="sxs-lookup"><span data-stu-id="97316-660">Attributes</span></span>| <span data-ttu-id="97316-661">説明</span><span class="sxs-lookup"><span data-stu-id="97316-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="97316-662">String</span><span class="sxs-lookup"><span data-stu-id="97316-662">String</span></span>||<span data-ttu-id="97316-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="97316-665">String</span><span class="sxs-lookup"><span data-stu-id="97316-665">String</span></span>||<span data-ttu-id="97316-666">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="97316-666">The subject of the item to be attached.</span></span> <span data-ttu-id="97316-667">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="97316-668">Object</span><span class="sxs-lookup"><span data-stu-id="97316-668">Object</span></span>| <span data-ttu-id="97316-669">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-669">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-670">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-671">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-671">Object</span></span>| <span data-ttu-id="97316-672">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-672">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-673">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97316-674">function</span><span class="sxs-lookup"><span data-stu-id="97316-674">function</span></span>| <span data-ttu-id="97316-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-675">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-676">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97316-677">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="97316-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="97316-678">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="97316-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97316-679">エラー</span><span class="sxs-lookup"><span data-stu-id="97316-679">Errors</span></span>

| <span data-ttu-id="97316-680">エラー コード</span><span class="sxs-lookup"><span data-stu-id="97316-680">Error code</span></span> | <span data-ttu-id="97316-681">説明</span><span class="sxs-lookup"><span data-stu-id="97316-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="97316-682">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="97316-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-683">要件</span><span class="sxs-lookup"><span data-stu-id="97316-683">Requirements</span></span>

|<span data-ttu-id="97316-684">要件</span><span class="sxs-lookup"><span data-stu-id="97316-684">Requirement</span></span>| <span data-ttu-id="97316-685">値</span><span class="sxs-lookup"><span data-stu-id="97316-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-686">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-687">1.1</span><span class="sxs-lookup"><span data-stu-id="97316-687">1.1</span></span>|
|[<span data-ttu-id="97316-688">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-690">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-691">作成</span><span class="sxs-lookup"><span data-stu-id="97316-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-692">例</span><span class="sxs-lookup"><span data-stu-id="97316-692">Example</span></span>

<span data-ttu-id="97316-693">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="97316-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="97316-694">close()</span><span class="sxs-lookup"><span data-stu-id="97316-694">close()</span></span>

<span data-ttu-id="97316-695">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="97316-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="97316-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="97316-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-698">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="97316-699">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="97316-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-700">要件</span><span class="sxs-lookup"><span data-stu-id="97316-700">Requirements</span></span>

|<span data-ttu-id="97316-701">要件</span><span class="sxs-lookup"><span data-stu-id="97316-701">Requirement</span></span>| <span data-ttu-id="97316-702">値</span><span class="sxs-lookup"><span data-stu-id="97316-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-703">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-704">1.3</span><span class="sxs-lookup"><span data-stu-id="97316-704">1.3</span></span>|
|[<span data-ttu-id="97316-705">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-706">制限あり</span><span class="sxs-lookup"><span data-stu-id="97316-706">Restricted</span></span>|
|[<span data-ttu-id="97316-707">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-708">新規作成</span><span class="sxs-lookup"><span data-stu-id="97316-708">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="97316-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="97316-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="97316-710">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-711">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="97316-712">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="97316-713">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="97316-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="97316-714">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="97316-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="97316-715">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="97316-716">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="97316-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-717">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-717">Parameters</span></span>

|<span data-ttu-id="97316-718">名前</span><span class="sxs-lookup"><span data-stu-id="97316-718">Name</span></span>| <span data-ttu-id="97316-719">型</span><span class="sxs-lookup"><span data-stu-id="97316-719">Type</span></span>| <span data-ttu-id="97316-720">説明</span><span class="sxs-lookup"><span data-stu-id="97316-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="97316-721">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="97316-721">String &#124; Object</span></span>| |<span data-ttu-id="97316-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="97316-724">**または**</span><span class="sxs-lookup"><span data-stu-id="97316-724">**OR**</span></span><br/><span data-ttu-id="97316-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="97316-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="97316-727">String</span><span class="sxs-lookup"><span data-stu-id="97316-727">String</span></span> | <span data-ttu-id="97316-728">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-728">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="97316-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="97316-732">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-732">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-733">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="97316-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="97316-734">String</span><span class="sxs-lookup"><span data-stu-id="97316-734">String</span></span> | | <span data-ttu-id="97316-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="97316-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="97316-737">String</span><span class="sxs-lookup"><span data-stu-id="97316-737">String</span></span> | | <span data-ttu-id="97316-738">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="97316-739">文字列</span><span class="sxs-lookup"><span data-stu-id="97316-739">String</span></span> | | <span data-ttu-id="97316-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="97316-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="97316-742">String</span><span class="sxs-lookup"><span data-stu-id="97316-742">String</span></span> | | <span data-ttu-id="97316-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="97316-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="97316-746">function</span><span class="sxs-lookup"><span data-stu-id="97316-746">function</span></span> | <span data-ttu-id="97316-747">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-747">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-748">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-749">要件</span><span class="sxs-lookup"><span data-stu-id="97316-749">Requirements</span></span>

|<span data-ttu-id="97316-750">要件</span><span class="sxs-lookup"><span data-stu-id="97316-750">Requirement</span></span>| <span data-ttu-id="97316-751">値</span><span class="sxs-lookup"><span data-stu-id="97316-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-752">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-753">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-753">1.0</span></span>|
|[<span data-ttu-id="97316-754">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-755">ReadItem</span></span>|
|[<span data-ttu-id="97316-756">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-757">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="97316-758">例</span><span class="sxs-lookup"><span data-stu-id="97316-758">Examples</span></span>

<span data-ttu-id="97316-759">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="97316-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="97316-760">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-760">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="97316-761">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-761">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="97316-762">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="97316-763">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="97316-764">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="97316-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="97316-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="97316-766">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-767">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="97316-768">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="97316-769">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="97316-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="97316-770">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="97316-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="97316-771">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="97316-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="97316-772">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="97316-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-773">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-773">Parameters</span></span>

|<span data-ttu-id="97316-774">名前</span><span class="sxs-lookup"><span data-stu-id="97316-774">Name</span></span>| <span data-ttu-id="97316-775">型</span><span class="sxs-lookup"><span data-stu-id="97316-775">Type</span></span>| <span data-ttu-id="97316-776">説明</span><span class="sxs-lookup"><span data-stu-id="97316-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="97316-777">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="97316-777">String &#124; Object</span></span>| | <span data-ttu-id="97316-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="97316-780">**または**</span><span class="sxs-lookup"><span data-stu-id="97316-780">**OR**</span></span><br/><span data-ttu-id="97316-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="97316-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="97316-783">String</span><span class="sxs-lookup"><span data-stu-id="97316-783">String</span></span> | <span data-ttu-id="97316-784">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-784">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="97316-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="97316-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="97316-788">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-788">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-789">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="97316-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="97316-790">String</span><span class="sxs-lookup"><span data-stu-id="97316-790">String</span></span> | | <span data-ttu-id="97316-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="97316-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="97316-793">String</span><span class="sxs-lookup"><span data-stu-id="97316-793">String</span></span> | | <span data-ttu-id="97316-794">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="97316-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="97316-795">文字列</span><span class="sxs-lookup"><span data-stu-id="97316-795">String</span></span> | | <span data-ttu-id="97316-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="97316-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="97316-798">String</span><span class="sxs-lookup"><span data-stu-id="97316-798">String</span></span> | | <span data-ttu-id="97316-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="97316-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="97316-802">function</span><span class="sxs-lookup"><span data-stu-id="97316-802">function</span></span> | <span data-ttu-id="97316-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-803">&lt;optional&gt;</span></span> | <span data-ttu-id="97316-804">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-805">要件</span><span class="sxs-lookup"><span data-stu-id="97316-805">Requirements</span></span>

|<span data-ttu-id="97316-806">要件</span><span class="sxs-lookup"><span data-stu-id="97316-806">Requirement</span></span>| <span data-ttu-id="97316-807">値</span><span class="sxs-lookup"><span data-stu-id="97316-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-808">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-809">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-809">1.0</span></span>|
|[<span data-ttu-id="97316-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-811">ReadItem</span></span>|
|[<span data-ttu-id="97316-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-813">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="97316-814">例</span><span class="sxs-lookup"><span data-stu-id="97316-814">Examples</span></span>

<span data-ttu-id="97316-815">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="97316-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="97316-816">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-816">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="97316-817">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-817">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="97316-818">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="97316-819">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="97316-820">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="97316-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="97316-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="97316-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="97316-822">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-823">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-824">要件</span><span class="sxs-lookup"><span data-stu-id="97316-824">Requirements</span></span>

|<span data-ttu-id="97316-825">要件</span><span class="sxs-lookup"><span data-stu-id="97316-825">Requirement</span></span>| <span data-ttu-id="97316-826">値</span><span class="sxs-lookup"><span data-stu-id="97316-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-827">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-828">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-828">1.0</span></span>|
|[<span data-ttu-id="97316-829">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-830">ReadItem</span></span>|
|[<span data-ttu-id="97316-831">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-832">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-833">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-833">Returns:</span></span>

<span data-ttu-id="97316-834">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="97316-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="97316-835">例</span><span class="sxs-lookup"><span data-stu-id="97316-835">Example</span></span>

<span data-ttu-id="97316-836">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="97316-836">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="97316-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="97316-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="97316-838">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-839">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-840">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-840">Parameters</span></span>

|<span data-ttu-id="97316-841">名前</span><span class="sxs-lookup"><span data-stu-id="97316-841">Name</span></span>| <span data-ttu-id="97316-842">型</span><span class="sxs-lookup"><span data-stu-id="97316-842">Type</span></span>| <span data-ttu-id="97316-843">説明</span><span class="sxs-lookup"><span data-stu-id="97316-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="97316-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="97316-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="97316-845">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="97316-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-846">Requirements</span><span class="sxs-lookup"><span data-stu-id="97316-846">Requirements</span></span>

|<span data-ttu-id="97316-847">要件</span><span class="sxs-lookup"><span data-stu-id="97316-847">Requirement</span></span>| <span data-ttu-id="97316-848">値</span><span class="sxs-lookup"><span data-stu-id="97316-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-849">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-850">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-850">1.0</span></span>|
|[<span data-ttu-id="97316-851">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-852">制限あり</span><span class="sxs-lookup"><span data-stu-id="97316-852">Restricted</span></span>|
|[<span data-ttu-id="97316-853">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-854">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-855">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-855">Returns:</span></span>

<span data-ttu-id="97316-856">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="97316-857">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="97316-858">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="97316-859">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="97316-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="97316-860">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="97316-860">Value of `entityType`</span></span> | <span data-ttu-id="97316-861">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="97316-861">Type of objects in returned array</span></span> | <span data-ttu-id="97316-862">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="97316-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="97316-863">String</span><span class="sxs-lookup"><span data-stu-id="97316-863">String</span></span> | <span data-ttu-id="97316-864">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="97316-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="97316-865">連絡先</span><span class="sxs-lookup"><span data-stu-id="97316-865">Contact</span></span> | <span data-ttu-id="97316-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97316-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="97316-867">文字列</span><span class="sxs-lookup"><span data-stu-id="97316-867">String</span></span> | <span data-ttu-id="97316-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97316-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="97316-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="97316-869">MeetingSuggestion</span></span> | <span data-ttu-id="97316-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97316-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="97316-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="97316-871">PhoneNumber</span></span> | <span data-ttu-id="97316-872">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="97316-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="97316-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="97316-873">TaskSuggestion</span></span> | <span data-ttu-id="97316-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="97316-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="97316-875">文字列</span><span class="sxs-lookup"><span data-stu-id="97316-875">String</span></span> | <span data-ttu-id="97316-876">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="97316-876">**Restricted**</span></span> |

<span data-ttu-id="97316-877">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="97316-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="97316-878">例</span><span class="sxs-lookup"><span data-stu-id="97316-878">Example</span></span>

<span data-ttu-id="97316-879">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="97316-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="97316-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="97316-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="97316-881">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-882">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="97316-883">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-884">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-884">Parameters</span></span>

|<span data-ttu-id="97316-885">名前</span><span class="sxs-lookup"><span data-stu-id="97316-885">Name</span></span>| <span data-ttu-id="97316-886">型</span><span class="sxs-lookup"><span data-stu-id="97316-886">Type</span></span>| <span data-ttu-id="97316-887">説明</span><span class="sxs-lookup"><span data-stu-id="97316-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="97316-888">String</span><span class="sxs-lookup"><span data-stu-id="97316-888">String</span></span>|<span data-ttu-id="97316-889">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="97316-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-890">要件</span><span class="sxs-lookup"><span data-stu-id="97316-890">Requirements</span></span>

|<span data-ttu-id="97316-891">要件</span><span class="sxs-lookup"><span data-stu-id="97316-891">Requirement</span></span>| <span data-ttu-id="97316-892">値</span><span class="sxs-lookup"><span data-stu-id="97316-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-893">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-894">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-894">1.0</span></span>|
|[<span data-ttu-id="97316-895">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-896">ReadItem</span></span>|
|[<span data-ttu-id="97316-897">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-898">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-899">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-899">Returns:</span></span>

<span data-ttu-id="97316-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="97316-902">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="97316-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="97316-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="97316-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="97316-904">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-905">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="97316-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="97316-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="97316-909">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="97316-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="97316-910">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="97316-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="97316-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="97316-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="97316-914">要件</span><span class="sxs-lookup"><span data-stu-id="97316-914">Requirements</span></span>

|<span data-ttu-id="97316-915">要件</span><span class="sxs-lookup"><span data-stu-id="97316-915">Requirement</span></span>| <span data-ttu-id="97316-916">値</span><span class="sxs-lookup"><span data-stu-id="97316-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-917">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-918">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-918">1.0</span></span>|
|[<span data-ttu-id="97316-919">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-920">ReadItem</span></span>|
|[<span data-ttu-id="97316-921">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-922">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-923">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-923">Returns:</span></span>

<span data-ttu-id="97316-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="97316-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="97316-926">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="97316-926">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="97316-927">Object</span><span class="sxs-lookup"><span data-stu-id="97316-927">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="97316-928">例</span><span class="sxs-lookup"><span data-stu-id="97316-928">Example</span></span>

<span data-ttu-id="97316-929">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="97316-929">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="97316-930">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="97316-930">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="97316-931">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-931">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-932">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-932">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="97316-933">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="97316-933">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="97316-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="97316-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-936">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-936">Parameters</span></span>

|<span data-ttu-id="97316-937">名前</span><span class="sxs-lookup"><span data-stu-id="97316-937">Name</span></span>| <span data-ttu-id="97316-938">型</span><span class="sxs-lookup"><span data-stu-id="97316-938">Type</span></span>| <span data-ttu-id="97316-939">説明</span><span class="sxs-lookup"><span data-stu-id="97316-939">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="97316-940">String</span><span class="sxs-lookup"><span data-stu-id="97316-940">String</span></span>|<span data-ttu-id="97316-941">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="97316-941">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-942">要件</span><span class="sxs-lookup"><span data-stu-id="97316-942">Requirements</span></span>

|<span data-ttu-id="97316-943">要件</span><span class="sxs-lookup"><span data-stu-id="97316-943">Requirement</span></span>| <span data-ttu-id="97316-944">値</span><span class="sxs-lookup"><span data-stu-id="97316-944">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-945">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-945">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-946">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-946">1.0</span></span>|
|[<span data-ttu-id="97316-947">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-947">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-948">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-948">ReadItem</span></span>|
|[<span data-ttu-id="97316-949">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-949">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-950">読み取り</span><span class="sxs-lookup"><span data-stu-id="97316-950">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-951">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-951">Returns:</span></span>

<span data-ttu-id="97316-952">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="97316-952">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="97316-953">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="97316-953">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="97316-954">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="97316-954">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="97316-955">例</span><span class="sxs-lookup"><span data-stu-id="97316-955">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="97316-956">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="97316-956">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="97316-957">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-957">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="97316-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="97316-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-960">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-960">Parameters</span></span>

|<span data-ttu-id="97316-961">名前</span><span class="sxs-lookup"><span data-stu-id="97316-961">Name</span></span>| <span data-ttu-id="97316-962">型</span><span class="sxs-lookup"><span data-stu-id="97316-962">Type</span></span>| <span data-ttu-id="97316-963">属性</span><span class="sxs-lookup"><span data-stu-id="97316-963">Attributes</span></span>| <span data-ttu-id="97316-964">説明</span><span class="sxs-lookup"><span data-stu-id="97316-964">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="97316-965">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="97316-965">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="97316-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="97316-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="97316-969">Object</span><span class="sxs-lookup"><span data-stu-id="97316-969">Object</span></span>| <span data-ttu-id="97316-970">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-970">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-971">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-971">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-972">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-972">Object</span></span>| <span data-ttu-id="97316-973">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-973">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-974">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-974">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97316-975">function</span><span class="sxs-lookup"><span data-stu-id="97316-975">function</span></span>||<span data-ttu-id="97316-976">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-976">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="97316-977">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="97316-977">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="97316-978">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="97316-978">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-979">要件</span><span class="sxs-lookup"><span data-stu-id="97316-979">Requirements</span></span>

|<span data-ttu-id="97316-980">要件</span><span class="sxs-lookup"><span data-stu-id="97316-980">Requirement</span></span>| <span data-ttu-id="97316-981">値</span><span class="sxs-lookup"><span data-stu-id="97316-981">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-982">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-982">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-983">1.2</span><span class="sxs-lookup"><span data-stu-id="97316-983">1.2</span></span>|
|[<span data-ttu-id="97316-984">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-984">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-985">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-985">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-986">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-986">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-987">作成</span><span class="sxs-lookup"><span data-stu-id="97316-987">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="97316-988">戻り値:</span><span class="sxs-lookup"><span data-stu-id="97316-988">Returns:</span></span>

<span data-ttu-id="97316-989">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="97316-989">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="97316-990">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="97316-990">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="97316-991">String</span><span class="sxs-lookup"><span data-stu-id="97316-991">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="97316-992">例</span><span class="sxs-lookup"><span data-stu-id="97316-992">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="97316-993">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="97316-993">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="97316-994">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="97316-994">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="97316-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="97316-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-998">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-998">Parameters</span></span>

|<span data-ttu-id="97316-999">名前</span><span class="sxs-lookup"><span data-stu-id="97316-999">Name</span></span>| <span data-ttu-id="97316-1000">型</span><span class="sxs-lookup"><span data-stu-id="97316-1000">Type</span></span>| <span data-ttu-id="97316-1001">属性</span><span class="sxs-lookup"><span data-stu-id="97316-1001">Attributes</span></span>| <span data-ttu-id="97316-1002">説明</span><span class="sxs-lookup"><span data-stu-id="97316-1002">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="97316-1003">function</span><span class="sxs-lookup"><span data-stu-id="97316-1003">function</span></span>||<span data-ttu-id="97316-1004">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1004">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="97316-1005">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1005">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="97316-1006">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="97316-1006">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="97316-1007">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1007">Object</span></span>| <span data-ttu-id="97316-1008">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1009">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="97316-1009">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="97316-1010">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="97316-1010">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-1011">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1011">Requirements</span></span>

|<span data-ttu-id="97316-1012">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1012">Requirement</span></span>| <span data-ttu-id="97316-1013">値</span><span class="sxs-lookup"><span data-stu-id="97316-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-1014">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-1015">1.0</span><span class="sxs-lookup"><span data-stu-id="97316-1015">1.0</span></span>|
|[<span data-ttu-id="97316-1016">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-1017">ReadItem</span><span class="sxs-lookup"><span data-stu-id="97316-1017">ReadItem</span></span>|
|[<span data-ttu-id="97316-1018">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-1019">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="97316-1019">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-1020">例</span><span class="sxs-lookup"><span data-stu-id="97316-1020">Example</span></span>

<span data-ttu-id="97316-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="97316-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="97316-1024">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="97316-1024">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="97316-1025">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="97316-1025">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="97316-1026">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="97316-1026">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="97316-1027">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="97316-1027">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="97316-1028">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="97316-1028">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="97316-1029">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="97316-1029">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-1030">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-1030">Parameters</span></span>

|<span data-ttu-id="97316-1031">名前</span><span class="sxs-lookup"><span data-stu-id="97316-1031">Name</span></span>| <span data-ttu-id="97316-1032">型</span><span class="sxs-lookup"><span data-stu-id="97316-1032">Type</span></span>| <span data-ttu-id="97316-1033">属性</span><span class="sxs-lookup"><span data-stu-id="97316-1033">Attributes</span></span>| <span data-ttu-id="97316-1034">説明</span><span class="sxs-lookup"><span data-stu-id="97316-1034">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="97316-1035">String</span><span class="sxs-lookup"><span data-stu-id="97316-1035">String</span></span>||<span data-ttu-id="97316-1036">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="97316-1036">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="97316-1037">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1037">Object</span></span>| <span data-ttu-id="97316-1038">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1039">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-1040">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1040">Object</span></span>| <span data-ttu-id="97316-1041">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1042">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="97316-1043">function</span><span class="sxs-lookup"><span data-stu-id="97316-1043">function</span></span>| <span data-ttu-id="97316-1044">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1045">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1045">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="97316-1046">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="97316-1046">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="97316-1047">エラー</span><span class="sxs-lookup"><span data-stu-id="97316-1047">Errors</span></span>

| <span data-ttu-id="97316-1048">エラー コード</span><span class="sxs-lookup"><span data-stu-id="97316-1048">Error code</span></span> | <span data-ttu-id="97316-1049">説明</span><span class="sxs-lookup"><span data-stu-id="97316-1049">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="97316-1050">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="97316-1050">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-1051">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1051">Requirements</span></span>

|<span data-ttu-id="97316-1052">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1052">Requirement</span></span>| <span data-ttu-id="97316-1053">値</span><span class="sxs-lookup"><span data-stu-id="97316-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-1054">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-1055">1.1</span><span class="sxs-lookup"><span data-stu-id="97316-1055">1.1</span></span>|
|[<span data-ttu-id="97316-1056">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-1058">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-1059">作成</span><span class="sxs-lookup"><span data-stu-id="97316-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-1060">例</span><span class="sxs-lookup"><span data-stu-id="97316-1060">Example</span></span>

<span data-ttu-id="97316-1061">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="97316-1061">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="97316-1062">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="97316-1062">saveAsync([options], callback)</span></span>

<span data-ttu-id="97316-1063">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="97316-1063">Asynchronously saves an item.</span></span>

<span data-ttu-id="97316-1064">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="97316-1064">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="97316-1065">Outlook on the web または online モードの Outlook では、アイテムはサーバーに保存されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1065">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="97316-1066">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1066">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-1067">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="97316-1067">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="97316-1068">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1068">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="97316-p168">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="97316-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="97316-1072">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="97316-1072">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="97316-1073">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="97316-1073">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="97316-1074">新規`saveAsync`作成モードで会議から呼び出された場合、メソッドは失敗します。</span><span class="sxs-lookup"><span data-stu-id="97316-1074">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="97316-1075">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="97316-1075">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="97316-1076">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1076">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-1077">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-1077">Parameters</span></span>

|<span data-ttu-id="97316-1078">名前</span><span class="sxs-lookup"><span data-stu-id="97316-1078">Name</span></span>| <span data-ttu-id="97316-1079">型</span><span class="sxs-lookup"><span data-stu-id="97316-1079">Type</span></span>| <span data-ttu-id="97316-1080">属性</span><span class="sxs-lookup"><span data-stu-id="97316-1080">Attributes</span></span>| <span data-ttu-id="97316-1081">説明</span><span class="sxs-lookup"><span data-stu-id="97316-1081">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="97316-1082">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1082">Object</span></span>| <span data-ttu-id="97316-1083">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1084">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-1085">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1085">Object</span></span>| <span data-ttu-id="97316-1086">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1087">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-1087">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="97316-1088">関数</span><span class="sxs-lookup"><span data-stu-id="97316-1088">function</span></span>||<span data-ttu-id="97316-1089">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="97316-1090">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1090">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="97316-1091">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1091">Requirements</span></span>

|<span data-ttu-id="97316-1092">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1092">Requirement</span></span>| <span data-ttu-id="97316-1093">値</span><span class="sxs-lookup"><span data-stu-id="97316-1093">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-1094">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-1094">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-1095">1.3</span><span class="sxs-lookup"><span data-stu-id="97316-1095">1.3</span></span>|
|[<span data-ttu-id="97316-1096">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-1096">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-1097">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-1097">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-1098">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-1098">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-1099">作成</span><span class="sxs-lookup"><span data-stu-id="97316-1099">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="97316-1100">例</span><span class="sxs-lookup"><span data-stu-id="97316-1100">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="97316-p170">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="97316-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="97316-1103">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="97316-1103">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="97316-1104">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="97316-1104">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="97316-p171">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="97316-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="97316-1108">パラメーター</span><span class="sxs-lookup"><span data-stu-id="97316-1108">Parameters</span></span>

|<span data-ttu-id="97316-1109">名前</span><span class="sxs-lookup"><span data-stu-id="97316-1109">Name</span></span>| <span data-ttu-id="97316-1110">型</span><span class="sxs-lookup"><span data-stu-id="97316-1110">Type</span></span>| <span data-ttu-id="97316-1111">属性</span><span class="sxs-lookup"><span data-stu-id="97316-1111">Attributes</span></span>| <span data-ttu-id="97316-1112">説明</span><span class="sxs-lookup"><span data-stu-id="97316-1112">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="97316-1113">String</span><span class="sxs-lookup"><span data-stu-id="97316-1113">String</span></span>||<span data-ttu-id="97316-p172">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="97316-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="97316-1117">Object</span><span class="sxs-lookup"><span data-stu-id="97316-1117">Object</span></span>| <span data-ttu-id="97316-1118">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1119">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="97316-1119">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="97316-1120">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="97316-1120">Object</span></span>| <span data-ttu-id="97316-1121">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1122">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="97316-1122">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="97316-1123">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="97316-1123">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="97316-1124">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="97316-1124">&lt;optional&gt;</span></span>|<span data-ttu-id="97316-1125">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1125">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="97316-1126">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1126">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="97316-1127">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1127">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="97316-1128">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1128">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="97316-1129">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1129">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="97316-1130">function</span><span class="sxs-lookup"><span data-stu-id="97316-1130">function</span></span>||<span data-ttu-id="97316-1131">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="97316-1131">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="97316-1132">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1132">Requirements</span></span>

|<span data-ttu-id="97316-1133">要件</span><span class="sxs-lookup"><span data-stu-id="97316-1133">Requirement</span></span>| <span data-ttu-id="97316-1134">値</span><span class="sxs-lookup"><span data-stu-id="97316-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="97316-1135">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="97316-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="97316-1136">1.2</span><span class="sxs-lookup"><span data-stu-id="97316-1136">1.2</span></span>|
|[<span data-ttu-id="97316-1137">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="97316-1137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="97316-1138">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="97316-1138">ReadWriteItem</span></span>|
|[<span data-ttu-id="97316-1139">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="97316-1139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="97316-1140">作成</span><span class="sxs-lookup"><span data-stu-id="97316-1140">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="97316-1141">例</span><span class="sxs-lookup"><span data-stu-id="97316-1141">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
