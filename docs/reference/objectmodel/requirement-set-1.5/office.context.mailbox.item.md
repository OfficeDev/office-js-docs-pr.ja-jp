---
title: Office.context.mailbox.item - requirement set 1.5
description: ''
ms.date: 09/23/2019
localization_priority: Priority
ms.openlocfilehash: a76440c2353483d0754aa0ba2fed98f051bc1c6f
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167369"
---
# <a name="item"></a><span data-ttu-id="d1469-102">item</span><span class="sxs-lookup"><span data-stu-id="d1469-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d1469-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d1469-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d1469-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-106">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-106">Requirements</span></span>

|<span data-ttu-id="d1469-107">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-107">Requirement</span></span>| <span data-ttu-id="d1469-108">値</span><span class="sxs-lookup"><span data-stu-id="d1469-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-110">1.0</span></span>|
|[<span data-ttu-id="d1469-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1469-112">Restricted</span></span>|
|[<span data-ttu-id="d1469-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d1469-115">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-115">Members and methods</span></span>

| <span data-ttu-id="d1469-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-116">Member</span></span> | <span data-ttu-id="d1469-117">種類</span><span class="sxs-lookup"><span data-stu-id="d1469-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d1469-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d1469-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d1469-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-119">Member</span></span> |
| [<span data-ttu-id="d1469-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d1469-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d1469-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-121">Member</span></span> |
| [<span data-ttu-id="d1469-122">body</span><span class="sxs-lookup"><span data-stu-id="d1469-122">body</span></span>](#body-body) | <span data-ttu-id="d1469-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-123">Member</span></span> |
| [<span data-ttu-id="d1469-124">cc</span><span class="sxs-lookup"><span data-stu-id="d1469-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d1469-125">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-125">Member</span></span> |
| [<span data-ttu-id="d1469-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="d1469-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d1469-127">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-127">Member</span></span> |
| [<span data-ttu-id="d1469-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d1469-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d1469-129">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-129">Member</span></span> |
| [<span data-ttu-id="d1469-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d1469-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d1469-131">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-131">Member</span></span> |
| [<span data-ttu-id="d1469-132">end</span><span class="sxs-lookup"><span data-stu-id="d1469-132">end</span></span>](#end-datetime) | <span data-ttu-id="d1469-133">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-133">Member</span></span> |
| [<span data-ttu-id="d1469-134">from</span><span class="sxs-lookup"><span data-stu-id="d1469-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="d1469-135">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-135">Member</span></span> |
| [<span data-ttu-id="d1469-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d1469-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d1469-137">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-137">Member</span></span> |
| [<span data-ttu-id="d1469-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="d1469-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d1469-139">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-139">Member</span></span> |
| [<span data-ttu-id="d1469-140">itemId</span><span class="sxs-lookup"><span data-stu-id="d1469-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d1469-141">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-141">Member</span></span> |
| [<span data-ttu-id="d1469-142">itemType</span><span class="sxs-lookup"><span data-stu-id="d1469-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d1469-143">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-143">Member</span></span> |
| [<span data-ttu-id="d1469-144">location</span><span class="sxs-lookup"><span data-stu-id="d1469-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="d1469-145">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-145">Member</span></span> |
| [<span data-ttu-id="d1469-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d1469-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d1469-147">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-147">Member</span></span> |
| [<span data-ttu-id="d1469-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d1469-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d1469-149">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-149">Member</span></span> |
| [<span data-ttu-id="d1469-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d1469-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d1469-151">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-151">Member</span></span> |
| [<span data-ttu-id="d1469-152">organizer</span><span class="sxs-lookup"><span data-stu-id="d1469-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="d1469-153">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-153">Member</span></span> |
| [<span data-ttu-id="d1469-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d1469-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d1469-155">Member</span><span class="sxs-lookup"><span data-stu-id="d1469-155">Member</span></span> |
| [<span data-ttu-id="d1469-156">sender</span><span class="sxs-lookup"><span data-stu-id="d1469-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d1469-157">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-157">Member</span></span> |
| [<span data-ttu-id="d1469-158">start</span><span class="sxs-lookup"><span data-stu-id="d1469-158">start</span></span>](#start-datetime) | <span data-ttu-id="d1469-159">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-159">Member</span></span> |
| [<span data-ttu-id="d1469-160">subject</span><span class="sxs-lookup"><span data-stu-id="d1469-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d1469-161">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-161">Member</span></span> |
| [<span data-ttu-id="d1469-162">to</span><span class="sxs-lookup"><span data-stu-id="d1469-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d1469-163">メンバー</span><span class="sxs-lookup"><span data-stu-id="d1469-163">Member</span></span> |
| [<span data-ttu-id="d1469-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d1469-165">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-165">Method</span></span> |
| [<span data-ttu-id="d1469-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d1469-167">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-167">Method</span></span> |
| [<span data-ttu-id="d1469-168">close</span><span class="sxs-lookup"><span data-stu-id="d1469-168">close</span></span>](#close) | <span data-ttu-id="d1469-169">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-169">Method</span></span> |
| [<span data-ttu-id="d1469-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d1469-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d1469-171">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-171">Method</span></span> |
| [<span data-ttu-id="d1469-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d1469-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d1469-173">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-173">Method</span></span> |
| [<span data-ttu-id="d1469-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="d1469-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d1469-175">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-175">Method</span></span> |
| [<span data-ttu-id="d1469-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d1469-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d1469-177">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-177">Method</span></span> |
| [<span data-ttu-id="d1469-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d1469-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d1469-179">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-179">Method</span></span> |
| [<span data-ttu-id="d1469-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d1469-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d1469-181">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-181">Method</span></span> |
| [<span data-ttu-id="d1469-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d1469-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d1469-183">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-183">Method</span></span> |
| [<span data-ttu-id="d1469-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d1469-185">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-185">Method</span></span> |
| [<span data-ttu-id="d1469-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d1469-187">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-187">Method</span></span> |
| [<span data-ttu-id="d1469-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d1469-189">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-189">Method</span></span> |
| [<span data-ttu-id="d1469-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d1469-191">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-191">Method</span></span> |
| [<span data-ttu-id="d1469-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d1469-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d1469-193">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d1469-194">例</span><span class="sxs-lookup"><span data-stu-id="d1469-194">Example</span></span>

<span data-ttu-id="d1469-195">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="d1469-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d1469-196">Members</span><span class="sxs-lookup"><span data-stu-id="d1469-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="d1469-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="d1469-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="d1469-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-200">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="d1469-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d1469-201">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1469-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-202">型</span><span class="sxs-lookup"><span data-stu-id="d1469-202">Type</span></span>

*   <span data-ttu-id="d1469-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="d1469-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-204">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-204">Requirements</span></span>

|<span data-ttu-id="d1469-205">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-205">Requirement</span></span>| <span data-ttu-id="d1469-206">値</span><span class="sxs-lookup"><span data-stu-id="d1469-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-207">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-208">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-208">1.0</span></span>|
|[<span data-ttu-id="d1469-209">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-210">ReadItem</span></span>|
|[<span data-ttu-id="d1469-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-212">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-213">例</span><span class="sxs-lookup"><span data-stu-id="d1469-213">Example</span></span>

<span data-ttu-id="d1469-214">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="d1469-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="d1469-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-215">bcc :[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-216">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d1469-217">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-218">型</span><span class="sxs-lookup"><span data-stu-id="d1469-218">Type</span></span>

*   [<span data-ttu-id="d1469-219">受信者</span><span class="sxs-lookup"><span data-stu-id="d1469-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-220">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-220">Requirements</span></span>

|<span data-ttu-id="d1469-221">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-221">Requirement</span></span>| <span data-ttu-id="d1469-222">値</span><span class="sxs-lookup"><span data-stu-id="d1469-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-224">1.1</span><span class="sxs-lookup"><span data-stu-id="d1469-224">1.1</span></span>|
|[<span data-ttu-id="d1469-225">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-226">ReadItem</span></span>|
|[<span data-ttu-id="d1469-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-228">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-229">例</span><span class="sxs-lookup"><span data-stu-id="d1469-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="d1469-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-230">body :[Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-231">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-232">型</span><span class="sxs-lookup"><span data-stu-id="d1469-232">Type</span></span>

*   [<span data-ttu-id="d1469-233">Body</span><span class="sxs-lookup"><span data-stu-id="d1469-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-234">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-234">Requirements</span></span>

|<span data-ttu-id="d1469-235">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-235">Requirement</span></span>| <span data-ttu-id="d1469-236">値</span><span class="sxs-lookup"><span data-stu-id="d1469-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-238">1.1</span><span class="sxs-lookup"><span data-stu-id="d1469-238">1.1</span></span>|
|[<span data-ttu-id="d1469-239">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-240">ReadItem</span></span>|
|[<span data-ttu-id="d1469-241">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-242">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-243">例</span><span class="sxs-lookup"><span data-stu-id="d1469-243">Example</span></span>

<span data-ttu-id="d1469-244">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d1469-245">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="d1469-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="d1469-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-247">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1469-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d1469-248">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-249">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-249">Read mode</span></span>

<span data-ttu-id="d1469-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-252">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-252">Compose mode</span></span>

<span data-ttu-id="d1469-253">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d1469-254">型</span><span class="sxs-lookup"><span data-stu-id="d1469-254">Type</span></span>

*   <span data-ttu-id="d1469-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-256">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-256">Requirements</span></span>

|<span data-ttu-id="d1469-257">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-257">Requirement</span></span>| <span data-ttu-id="d1469-258">値</span><span class="sxs-lookup"><span data-stu-id="d1469-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-260">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-260">1.0</span></span>|
|[<span data-ttu-id="d1469-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-262">ReadItem</span></span>|
|[<span data-ttu-id="d1469-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d1469-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d1469-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="d1469-266">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d1469-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="d1469-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d1469-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-271">型</span><span class="sxs-lookup"><span data-stu-id="d1469-271">Type</span></span>

*   <span data-ttu-id="d1469-272">String</span><span class="sxs-lookup"><span data-stu-id="d1469-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-273">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-273">Requirements</span></span>

|<span data-ttu-id="d1469-274">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-274">Requirement</span></span>| <span data-ttu-id="d1469-275">値</span><span class="sxs-lookup"><span data-stu-id="d1469-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-277">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-277">1.0</span></span>|
|[<span data-ttu-id="d1469-278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-279">ReadItem</span></span>|
|[<span data-ttu-id="d1469-280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-281">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-282">例</span><span class="sxs-lookup"><span data-stu-id="d1469-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d1469-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d1469-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="d1469-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-286">型</span><span class="sxs-lookup"><span data-stu-id="d1469-286">Type</span></span>

*   <span data-ttu-id="d1469-287">日付</span><span class="sxs-lookup"><span data-stu-id="d1469-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-288">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-288">Requirements</span></span>

|<span data-ttu-id="d1469-289">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-289">Requirement</span></span>| <span data-ttu-id="d1469-290">値</span><span class="sxs-lookup"><span data-stu-id="d1469-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-292">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-292">1.0</span></span>|
|[<span data-ttu-id="d1469-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-294">ReadItem</span></span>|
|[<span data-ttu-id="d1469-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-297">例</span><span class="sxs-lookup"><span data-stu-id="d1469-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d1469-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d1469-298">dateTimeModified :Date</span></span>

<span data-ttu-id="d1469-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-301">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-302">種類</span><span class="sxs-lookup"><span data-stu-id="d1469-302">Type</span></span>

*   <span data-ttu-id="d1469-303">日付</span><span class="sxs-lookup"><span data-stu-id="d1469-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-304">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-304">Requirements</span></span>

|<span data-ttu-id="d1469-305">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-305">Requirement</span></span>| <span data-ttu-id="d1469-306">値</span><span class="sxs-lookup"><span data-stu-id="d1469-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-308">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-308">1.0</span></span>|
|[<span data-ttu-id="d1469-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-310">ReadItem</span></span>|
|[<span data-ttu-id="d1469-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-313">例</span><span class="sxs-lookup"><span data-stu-id="d1469-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="d1469-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-314">end :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d1469-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-318">Read mode</span></span>

<span data-ttu-id="d1469-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-320">Compose mode</span></span>

<span data-ttu-id="d1469-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d1469-322">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1469-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d1469-323">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d1469-324">型</span><span class="sxs-lookup"><span data-stu-id="d1469-324">Type</span></span>

*   <span data-ttu-id="d1469-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-326">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-326">Requirements</span></span>

|<span data-ttu-id="d1469-327">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-327">Requirement</span></span>| <span data-ttu-id="d1469-328">値</span><span class="sxs-lookup"><span data-stu-id="d1469-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-330">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-330">1.0</span></span>|
|[<span data-ttu-id="d1469-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-332">ReadItem</span></span>|
|[<span data-ttu-id="d1469-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="d1469-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-335">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="d1469-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-340">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d1469-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-341">型</span><span class="sxs-lookup"><span data-stu-id="d1469-341">Type</span></span>

*   [<span data-ttu-id="d1469-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d1469-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-343">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-343">Requirements</span></span>

|<span data-ttu-id="d1469-344">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-344">Requirement</span></span>| <span data-ttu-id="d1469-345">値</span><span class="sxs-lookup"><span data-stu-id="d1469-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-347">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-347">1.0</span></span>|
|[<span data-ttu-id="d1469-348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-349">ReadItem</span></span>|
|[<span data-ttu-id="d1469-350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-351">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-352">例</span><span class="sxs-lookup"><span data-stu-id="d1469-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d1469-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d1469-353">internetMessageId :String</span></span>

<span data-ttu-id="d1469-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-356">型</span><span class="sxs-lookup"><span data-stu-id="d1469-356">Type</span></span>

*   <span data-ttu-id="d1469-357">String</span><span class="sxs-lookup"><span data-stu-id="d1469-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-358">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-358">Requirements</span></span>

|<span data-ttu-id="d1469-359">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-359">Requirement</span></span>| <span data-ttu-id="d1469-360">値</span><span class="sxs-lookup"><span data-stu-id="d1469-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-362">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-362">1.0</span></span>|
|[<span data-ttu-id="d1469-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-364">ReadItem</span></span>|
|[<span data-ttu-id="d1469-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-367">例</span><span class="sxs-lookup"><span data-stu-id="d1469-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d1469-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d1469-368">itemClass :String</span></span>

<span data-ttu-id="d1469-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d1469-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="d1469-373">型</span><span class="sxs-lookup"><span data-stu-id="d1469-373">Type</span></span> | <span data-ttu-id="d1469-374">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-374">Description</span></span> | <span data-ttu-id="d1469-375">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="d1469-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="d1469-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="d1469-376">Appointment items</span></span> | <span data-ttu-id="d1469-377">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1469-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="d1469-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="d1469-378">Message items</span></span> | <span data-ttu-id="d1469-379">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="d1469-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-381">型</span><span class="sxs-lookup"><span data-stu-id="d1469-381">Type</span></span>

*   <span data-ttu-id="d1469-382">String</span><span class="sxs-lookup"><span data-stu-id="d1469-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-383">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-383">Requirements</span></span>

|<span data-ttu-id="d1469-384">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-384">Requirement</span></span>| <span data-ttu-id="d1469-385">値</span><span class="sxs-lookup"><span data-stu-id="d1469-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-387">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-387">1.0</span></span>|
|[<span data-ttu-id="d1469-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-389">ReadItem</span></span>|
|[<span data-ttu-id="d1469-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-392">例</span><span class="sxs-lookup"><span data-stu-id="d1469-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d1469-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d1469-393">(nullable) itemId :String</span></span>

<span data-ttu-id="d1469-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="d1469-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d1469-397">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="d1469-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d1469-398">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1469-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d1469-399">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1469-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d1469-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-402">型</span><span class="sxs-lookup"><span data-stu-id="d1469-402">Type</span></span>

*   <span data-ttu-id="d1469-403">String</span><span class="sxs-lookup"><span data-stu-id="d1469-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-404">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-404">Requirements</span></span>

|<span data-ttu-id="d1469-405">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-405">Requirement</span></span>| <span data-ttu-id="d1469-406">値</span><span class="sxs-lookup"><span data-stu-id="d1469-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-407">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-408">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-408">1.0</span></span>|
|[<span data-ttu-id="d1469-409">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-410">ReadItem</span></span>|
|[<span data-ttu-id="d1469-411">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-412">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-413">例</span><span class="sxs-lookup"><span data-stu-id="d1469-413">Example</span></span>

<span data-ttu-id="d1469-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="d1469-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-417">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d1469-418">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="d1469-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-419">型</span><span class="sxs-lookup"><span data-stu-id="d1469-419">Type</span></span>

*   [<span data-ttu-id="d1469-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d1469-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-421">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-421">Requirements</span></span>

|<span data-ttu-id="d1469-422">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-422">Requirement</span></span>| <span data-ttu-id="d1469-423">値</span><span class="sxs-lookup"><span data-stu-id="d1469-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-425">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-425">1.0</span></span>|
|[<span data-ttu-id="d1469-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-427">ReadItem</span></span>|
|[<span data-ttu-id="d1469-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-429">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-430">例</span><span class="sxs-lookup"><span data-stu-id="d1469-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="d1469-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-431">location :String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-432">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-433">Read mode</span></span>

<span data-ttu-id="d1469-434">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-435">Compose mode</span></span>

<span data-ttu-id="d1469-436">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d1469-437">型</span><span class="sxs-lookup"><span data-stu-id="d1469-437">Type</span></span>

*   <span data-ttu-id="d1469-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-439">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-439">Requirements</span></span>

|<span data-ttu-id="d1469-440">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-440">Requirement</span></span>| <span data-ttu-id="d1469-441">値</span><span class="sxs-lookup"><span data-stu-id="d1469-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-443">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-443">1.0</span></span>|
|[<span data-ttu-id="d1469-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-445">ReadItem</span></span>|
|[<span data-ttu-id="d1469-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-447">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d1469-448">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d1469-448">normalizedSubject :String</span></span>

<span data-ttu-id="d1469-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d1469-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-453">型</span><span class="sxs-lookup"><span data-stu-id="d1469-453">Type</span></span>

*   <span data-ttu-id="d1469-454">String</span><span class="sxs-lookup"><span data-stu-id="d1469-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-455">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-455">Requirements</span></span>

|<span data-ttu-id="d1469-456">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-456">Requirement</span></span>| <span data-ttu-id="d1469-457">値</span><span class="sxs-lookup"><span data-stu-id="d1469-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-458">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-459">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-459">1.0</span></span>|
|[<span data-ttu-id="d1469-460">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-461">ReadItem</span></span>|
|[<span data-ttu-id="d1469-462">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-463">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-464">例</span><span class="sxs-lookup"><span data-stu-id="d1469-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="d1469-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-465">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-466">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-467">型</span><span class="sxs-lookup"><span data-stu-id="d1469-467">Type</span></span>

*   [<span data-ttu-id="d1469-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d1469-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-469">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-469">Requirements</span></span>

|<span data-ttu-id="d1469-470">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-470">Requirement</span></span>| <span data-ttu-id="d1469-471">値</span><span class="sxs-lookup"><span data-stu-id="d1469-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-472">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-473">1.3</span><span class="sxs-lookup"><span data-stu-id="d1469-473">1.3</span></span>|
|[<span data-ttu-id="d1469-474">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-475">ReadItem</span></span>|
|[<span data-ttu-id="d1469-476">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-477">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-478">例</span><span class="sxs-lookup"><span data-stu-id="d1469-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="d1469-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-480">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1469-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d1469-481">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-482">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-482">Read mode</span></span>

<span data-ttu-id="d1469-483">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-484">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-484">Compose mode</span></span>

<span data-ttu-id="d1469-485">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d1469-486">型</span><span class="sxs-lookup"><span data-stu-id="d1469-486">Type</span></span>

*   <span data-ttu-id="d1469-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-488">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-488">Requirements</span></span>

|<span data-ttu-id="d1469-489">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-489">Requirement</span></span>| <span data-ttu-id="d1469-490">値</span><span class="sxs-lookup"><span data-stu-id="d1469-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-491">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-492">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-492">1.0</span></span>|
|[<span data-ttu-id="d1469-493">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-494">ReadItem</span></span>|
|[<span data-ttu-id="d1469-495">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-496">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="d1469-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-497">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-500">型</span><span class="sxs-lookup"><span data-stu-id="d1469-500">Type</span></span>

*   [<span data-ttu-id="d1469-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d1469-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-502">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-502">Requirements</span></span>

|<span data-ttu-id="d1469-503">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-503">Requirement</span></span>| <span data-ttu-id="d1469-504">値</span><span class="sxs-lookup"><span data-stu-id="d1469-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-505">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-506">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-506">1.0</span></span>|
|[<span data-ttu-id="d1469-507">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-508">ReadItem</span></span>|
|[<span data-ttu-id="d1469-509">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-510">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-511">例</span><span class="sxs-lookup"><span data-stu-id="d1469-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="d1469-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-513">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1469-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d1469-514">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-515">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-515">Read mode</span></span>

<span data-ttu-id="d1469-516">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-517">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-517">Compose mode</span></span>

<span data-ttu-id="d1469-518">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d1469-519">型</span><span class="sxs-lookup"><span data-stu-id="d1469-519">Type</span></span>

*   <span data-ttu-id="d1469-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-521">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-521">Requirements</span></span>

|<span data-ttu-id="d1469-522">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-522">Requirement</span></span>| <span data-ttu-id="d1469-523">値</span><span class="sxs-lookup"><span data-stu-id="d1469-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-524">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-525">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-525">1.0</span></span>|
|[<span data-ttu-id="d1469-526">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-527">ReadItem</span></span>|
|[<span data-ttu-id="d1469-528">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-529">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="d1469-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-530">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="d1469-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d1469-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-535">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="d1469-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d1469-536">型</span><span class="sxs-lookup"><span data-stu-id="d1469-536">Type</span></span>

*   [<span data-ttu-id="d1469-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d1469-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="d1469-538">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-538">Requirements</span></span>

|<span data-ttu-id="d1469-539">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-539">Requirement</span></span>| <span data-ttu-id="d1469-540">値</span><span class="sxs-lookup"><span data-stu-id="d1469-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-541">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-542">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-542">1.0</span></span>|
|[<span data-ttu-id="d1469-543">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-544">ReadItem</span></span>|
|[<span data-ttu-id="d1469-545">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-546">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-547">例</span><span class="sxs-lookup"><span data-stu-id="d1469-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="d1469-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-548">start :Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-549">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d1469-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-552">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-552">Read mode</span></span>

<span data-ttu-id="d1469-553">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-554">Compose mode</span></span>

<span data-ttu-id="d1469-555">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d1469-556">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="d1469-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d1469-557">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d1469-558">型</span><span class="sxs-lookup"><span data-stu-id="d1469-558">Type</span></span>

*   <span data-ttu-id="d1469-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-560">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-560">Requirements</span></span>

|<span data-ttu-id="d1469-561">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-561">Requirement</span></span>| <span data-ttu-id="d1469-562">値</span><span class="sxs-lookup"><span data-stu-id="d1469-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-563">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-564">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-564">1.0</span></span>|
|[<span data-ttu-id="d1469-565">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-566">ReadItem</span></span>|
|[<span data-ttu-id="d1469-567">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-568">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="d1469-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-569">subject :String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-570">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d1469-571">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="d1469-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-572">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-572">Read mode</span></span>

<span data-ttu-id="d1469-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-575">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-575">Compose mode</span></span>

<span data-ttu-id="d1469-576">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d1469-577">型</span><span class="sxs-lookup"><span data-stu-id="d1469-577">Type</span></span>

*   <span data-ttu-id="d1469-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-579">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-579">Requirements</span></span>

|<span data-ttu-id="d1469-580">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-580">Requirement</span></span>| <span data-ttu-id="d1469-581">値</span><span class="sxs-lookup"><span data-stu-id="d1469-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-582">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-583">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-583">1.0</span></span>|
|[<span data-ttu-id="d1469-584">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-585">ReadItem</span></span>|
|[<span data-ttu-id="d1469-586">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-587">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="d1469-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="d1469-589">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d1469-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d1469-590">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d1469-591">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="d1469-591">Read mode</span></span>

<span data-ttu-id="d1469-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d1469-594">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="d1469-594">Compose mode</span></span>

<span data-ttu-id="d1469-595">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d1469-596">型</span><span class="sxs-lookup"><span data-stu-id="d1469-596">Type</span></span>

*   <span data-ttu-id="d1469-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-598">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-598">Requirements</span></span>

|<span data-ttu-id="d1469-599">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-599">Requirement</span></span>| <span data-ttu-id="d1469-600">値</span><span class="sxs-lookup"><span data-stu-id="d1469-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-601">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-602">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-602">1.0</span></span>|
|[<span data-ttu-id="d1469-603">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-604">ReadItem</span></span>|
|[<span data-ttu-id="d1469-605">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-606">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d1469-607">メソッド</span><span class="sxs-lookup"><span data-stu-id="d1469-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d1469-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d1469-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d1469-609">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d1469-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d1469-610">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="d1469-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d1469-611">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-612">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-612">Parameters</span></span>

|<span data-ttu-id="d1469-613">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-613">Name</span></span>| <span data-ttu-id="d1469-614">型</span><span class="sxs-lookup"><span data-stu-id="d1469-614">Type</span></span>| <span data-ttu-id="d1469-615">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-615">Attributes</span></span>| <span data-ttu-id="d1469-616">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="d1469-617">String</span><span class="sxs-lookup"><span data-stu-id="d1469-617">String</span></span>||<span data-ttu-id="d1469-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d1469-620">String</span><span class="sxs-lookup"><span data-stu-id="d1469-620">String</span></span>||<span data-ttu-id="d1469-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d1469-623">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-623">Object</span></span>| <span data-ttu-id="d1469-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-624">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-625">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="d1469-626">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-626">Object</span></span> | <span data-ttu-id="d1469-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-627">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-628">開発者は、コールバック メソッドでアクセスする任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="d1469-629">Boolean</span><span class="sxs-lookup"><span data-stu-id="d1469-629">Boolean</span></span> | <span data-ttu-id="d1469-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-630">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-631">`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="d1469-632">function</span><span class="sxs-lookup"><span data-stu-id="d1469-632">function</span></span>| <span data-ttu-id="d1469-633">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-633">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-634">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d1469-635">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d1469-636">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1469-637">エラー</span><span class="sxs-lookup"><span data-stu-id="d1469-637">Errors</span></span>

| <span data-ttu-id="d1469-638">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1469-638">Error code</span></span> | <span data-ttu-id="d1469-639">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="d1469-640">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="d1469-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="d1469-641">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="d1469-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d1469-642">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d1469-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-643">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-643">Requirements</span></span>

|<span data-ttu-id="d1469-644">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-644">Requirement</span></span>| <span data-ttu-id="d1469-645">値</span><span class="sxs-lookup"><span data-stu-id="d1469-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-646">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-647">1.1</span><span class="sxs-lookup"><span data-stu-id="d1469-647">1.1</span></span>|
|[<span data-ttu-id="d1469-648">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d1469-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="d1469-650">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-651">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d1469-652">例</span><span class="sxs-lookup"><span data-stu-id="d1469-652">Examples</span></span>

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

<span data-ttu-id="d1469-653">次の例では、インライン添付ファイルとしてイメージ ファイルを追加し、メッセージの本文の添付ファイルを参照します。</span><span class="sxs-lookup"><span data-stu-id="d1469-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d1469-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d1469-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d1469-655">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="d1469-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d1469-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d1469-659">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d1469-660">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="d1469-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-661">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-661">Parameters</span></span>

|<span data-ttu-id="d1469-662">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-662">Name</span></span>| <span data-ttu-id="d1469-663">型</span><span class="sxs-lookup"><span data-stu-id="d1469-663">Type</span></span>| <span data-ttu-id="d1469-664">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-664">Attributes</span></span>| <span data-ttu-id="d1469-665">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="d1469-666">String</span><span class="sxs-lookup"><span data-stu-id="d1469-666">String</span></span>||<span data-ttu-id="d1469-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="d1469-669">String</span><span class="sxs-lookup"><span data-stu-id="d1469-669">String</span></span>||<span data-ttu-id="d1469-670">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="d1469-670">The subject of the item to be attached.</span></span> <span data-ttu-id="d1469-671">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="d1469-672">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-672">Object</span></span>| <span data-ttu-id="d1469-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-673">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-674">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d1469-675">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-675">Object</span></span>| <span data-ttu-id="d1469-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-676">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-677">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d1469-678">function</span><span class="sxs-lookup"><span data-stu-id="d1469-678">function</span></span>| <span data-ttu-id="d1469-679">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-679">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-680">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d1469-681">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d1469-682">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1469-683">エラー</span><span class="sxs-lookup"><span data-stu-id="d1469-683">Errors</span></span>

| <span data-ttu-id="d1469-684">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1469-684">Error code</span></span> | <span data-ttu-id="d1469-685">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="d1469-686">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="d1469-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-687">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-687">Requirements</span></span>

|<span data-ttu-id="d1469-688">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-688">Requirement</span></span>| <span data-ttu-id="d1469-689">値</span><span class="sxs-lookup"><span data-stu-id="d1469-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-690">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-691">1.1</span><span class="sxs-lookup"><span data-stu-id="d1469-691">1.1</span></span>|
|[<span data-ttu-id="d1469-692">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d1469-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="d1469-694">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-695">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-696">例</span><span class="sxs-lookup"><span data-stu-id="d1469-696">Example</span></span>

<span data-ttu-id="d1469-697">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d1469-698">close()</span><span class="sxs-lookup"><span data-stu-id="d1469-698">close()</span></span>

<span data-ttu-id="d1469-699">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="d1469-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d1469-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-702">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d1469-703">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="d1469-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-704">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-704">Requirements</span></span>

|<span data-ttu-id="d1469-705">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-705">Requirement</span></span>| <span data-ttu-id="d1469-706">値</span><span class="sxs-lookup"><span data-stu-id="d1469-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-707">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-708">1.3</span><span class="sxs-lookup"><span data-stu-id="d1469-708">1.3</span></span>|
|[<span data-ttu-id="d1469-709">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-710">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1469-710">Restricted</span></span>|
|[<span data-ttu-id="d1469-711">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-712">新規作成</span><span class="sxs-lookup"><span data-stu-id="d1469-712">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d1469-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d1469-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d1469-714">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-715">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d1469-716">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d1469-717">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d1469-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d1469-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d1469-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-721">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-721">Parameters</span></span>

| <span data-ttu-id="d1469-722">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-722">Name</span></span> | <span data-ttu-id="d1469-723">型</span><span class="sxs-lookup"><span data-stu-id="d1469-723">Type</span></span> | <span data-ttu-id="d1469-724">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-724">Attributes</span></span> | <span data-ttu-id="d1469-725">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="d1469-726">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d1469-726">String &#124; Object</span></span>| |<span data-ttu-id="d1469-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d1469-729">**または**</span><span class="sxs-lookup"><span data-stu-id="d1469-729">**OR**</span></span><br/><span data-ttu-id="d1469-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1469-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d1469-732">String</span><span class="sxs-lookup"><span data-stu-id="d1469-732">String</span></span> | <span data-ttu-id="d1469-733">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-733">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d1469-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d1469-737">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-737">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-738">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d1469-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d1469-739">String</span><span class="sxs-lookup"><span data-stu-id="d1469-739">String</span></span> | | <span data-ttu-id="d1469-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d1469-742">String</span><span class="sxs-lookup"><span data-stu-id="d1469-742">String</span></span> | | <span data-ttu-id="d1469-743">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d1469-744">String</span><span class="sxs-lookup"><span data-stu-id="d1469-744">String</span></span> | | <span data-ttu-id="d1469-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d1469-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="d1469-747">ブール値</span><span class="sxs-lookup"><span data-stu-id="d1469-747">Boolean</span></span> | | <span data-ttu-id="d1469-p144">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d1469-750">String</span><span class="sxs-lookup"><span data-stu-id="d1469-750">String</span></span> | | <span data-ttu-id="d1469-p145">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d1469-754">function</span><span class="sxs-lookup"><span data-stu-id="d1469-754">function</span></span> | <span data-ttu-id="d1469-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-755">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-756">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-757">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-757">Requirements</span></span>

|<span data-ttu-id="d1469-758">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-758">Requirement</span></span>| <span data-ttu-id="d1469-759">値</span><span class="sxs-lookup"><span data-stu-id="d1469-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-760">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-761">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-761">1.0</span></span>|
|[<span data-ttu-id="d1469-762">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-763">ReadItem</span></span>|
|[<span data-ttu-id="d1469-764">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-765">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d1469-766">例</span><span class="sxs-lookup"><span data-stu-id="d1469-766">Examples</span></span>

<span data-ttu-id="d1469-767">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d1469-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d1469-768">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d1469-769">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d1469-770">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d1469-771">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d1469-772">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d1469-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d1469-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d1469-774">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-775">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d1469-776">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d1469-777">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="d1469-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d1469-p146">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook on the web とデスクトップ クライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。添付ファイルの追加に失敗すると、フォーム UI にエラーが表示されます。表示できない場合に、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="d1469-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-781">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-781">Parameters</span></span>

| <span data-ttu-id="d1469-782">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-782">Name</span></span> | <span data-ttu-id="d1469-783">型</span><span class="sxs-lookup"><span data-stu-id="d1469-783">Type</span></span> | <span data-ttu-id="d1469-784">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-784">Attributes</span></span> | <span data-ttu-id="d1469-785">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="d1469-786">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="d1469-786">String &#124; Object</span></span>| | <span data-ttu-id="d1469-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d1469-789">**または**</span><span class="sxs-lookup"><span data-stu-id="d1469-789">**OR**</span></span><br/><span data-ttu-id="d1469-p148">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="d1469-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="d1469-792">String</span><span class="sxs-lookup"><span data-stu-id="d1469-792">String</span></span> | <span data-ttu-id="d1469-793">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-793">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="d1469-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="d1469-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d1469-797">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-797">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-798">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="d1469-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="d1469-799">String</span><span class="sxs-lookup"><span data-stu-id="d1469-799">String</span></span> | | <span data-ttu-id="d1469-p150">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="d1469-802">String</span><span class="sxs-lookup"><span data-stu-id="d1469-802">String</span></span> | | <span data-ttu-id="d1469-803">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="d1469-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="d1469-804">String</span><span class="sxs-lookup"><span data-stu-id="d1469-804">String</span></span> | | <span data-ttu-id="d1469-p151">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="d1469-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="d1469-807">ブール値</span><span class="sxs-lookup"><span data-stu-id="d1469-807">Boolean</span></span> | | <span data-ttu-id="d1469-p152">`type` が `file` に設定されている場合にのみ使用されます。`true` の場合、添付ファイルがインラインでメッセージ本文に表示され、添付ファイル一覧に表示されないことを示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="d1469-810">String</span><span class="sxs-lookup"><span data-stu-id="d1469-810">String</span></span> | | <span data-ttu-id="d1469-p153">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="d1469-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="d1469-814">function</span><span class="sxs-lookup"><span data-stu-id="d1469-814">function</span></span> | <span data-ttu-id="d1469-815">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-815">&lt;optional&gt;</span></span> | <span data-ttu-id="d1469-816">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-817">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-817">Requirements</span></span>

|<span data-ttu-id="d1469-818">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-818">Requirement</span></span>| <span data-ttu-id="d1469-819">値</span><span class="sxs-lookup"><span data-stu-id="d1469-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-820">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-821">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-821">1.0</span></span>|
|[<span data-ttu-id="d1469-822">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-823">ReadItem</span></span>|
|[<span data-ttu-id="d1469-824">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-825">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d1469-826">例</span><span class="sxs-lookup"><span data-stu-id="d1469-826">Examples</span></span>

<span data-ttu-id="d1469-827">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="d1469-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d1469-828">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-828">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d1469-829">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-829">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d1469-830">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d1469-831">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d1469-832">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="d1469-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="d1469-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="d1469-833">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="d1469-834">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-835">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-836">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-836">Requirements</span></span>

|<span data-ttu-id="d1469-837">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-837">Requirement</span></span>| <span data-ttu-id="d1469-838">値</span><span class="sxs-lookup"><span data-stu-id="d1469-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-840">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-840">1.0</span></span>|
|[<span data-ttu-id="d1469-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-842">ReadItem</span></span>|
|[<span data-ttu-id="d1469-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-845">Returns:</span></span>

<span data-ttu-id="d1469-846">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="d1469-846">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="d1469-847">例</span><span class="sxs-lookup"><span data-stu-id="d1469-847">Example</span></span>

<span data-ttu-id="d1469-848">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="d1469-848">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="d1469-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="d1469-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="d1469-850">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-851">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-852">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-852">Parameters</span></span>

|<span data-ttu-id="d1469-853">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-853">Name</span></span>| <span data-ttu-id="d1469-854">種類</span><span class="sxs-lookup"><span data-stu-id="d1469-854">Type</span></span>| <span data-ttu-id="d1469-855">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="d1469-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d1469-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="d1469-857">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="d1469-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-858">Requirements</span><span class="sxs-lookup"><span data-stu-id="d1469-858">Requirements</span></span>

|<span data-ttu-id="d1469-859">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-859">Requirement</span></span>| <span data-ttu-id="d1469-860">値</span><span class="sxs-lookup"><span data-stu-id="d1469-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-861">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-862">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-862">1.0</span></span>|
|[<span data-ttu-id="d1469-863">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-864">制限あり</span><span class="sxs-lookup"><span data-stu-id="d1469-864">Restricted</span></span>|
|[<span data-ttu-id="d1469-865">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-866">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-867">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-867">Returns:</span></span>

<span data-ttu-id="d1469-868">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d1469-869">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d1469-870">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d1469-871">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="d1469-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="d1469-872">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="d1469-872">Value of `entityType`</span></span> | <span data-ttu-id="d1469-873">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="d1469-873">Type of objects in returned array</span></span> | <span data-ttu-id="d1469-874">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="d1469-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="d1469-875">文字列</span><span class="sxs-lookup"><span data-stu-id="d1469-875">String</span></span> | <span data-ttu-id="d1469-876">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d1469-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="d1469-877">連絡先</span><span class="sxs-lookup"><span data-stu-id="d1469-877">Contact</span></span> | <span data-ttu-id="d1469-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d1469-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="d1469-879">文字列</span><span class="sxs-lookup"><span data-stu-id="d1469-879">String</span></span> | <span data-ttu-id="d1469-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d1469-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="d1469-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d1469-881">MeetingSuggestion</span></span> | <span data-ttu-id="d1469-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d1469-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="d1469-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d1469-883">PhoneNumber</span></span> | <span data-ttu-id="d1469-884">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d1469-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="d1469-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d1469-885">TaskSuggestion</span></span> | <span data-ttu-id="d1469-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d1469-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="d1469-887">文字列</span><span class="sxs-lookup"><span data-stu-id="d1469-887">String</span></span> | <span data-ttu-id="d1469-888">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="d1469-888">**Restricted**</span></span> |

<span data-ttu-id="d1469-889">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="d1469-889">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="d1469-890">例</span><span class="sxs-lookup"><span data-stu-id="d1469-890">Example</span></span>

<span data-ttu-id="d1469-891">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="d1469-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="d1469-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="d1469-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="d1469-893">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-894">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d1469-895">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-896">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-896">Parameters</span></span>

|<span data-ttu-id="d1469-897">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-897">Name</span></span>| <span data-ttu-id="d1469-898">種類</span><span class="sxs-lookup"><span data-stu-id="d1469-898">Type</span></span>| <span data-ttu-id="d1469-899">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d1469-900">String</span><span class="sxs-lookup"><span data-stu-id="d1469-900">String</span></span>|<span data-ttu-id="d1469-901">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d1469-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-902">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-902">Requirements</span></span>

|<span data-ttu-id="d1469-903">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-903">Requirement</span></span>| <span data-ttu-id="d1469-904">値</span><span class="sxs-lookup"><span data-stu-id="d1469-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-905">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-906">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-906">1.0</span></span>|
|[<span data-ttu-id="d1469-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-908">ReadItem</span></span>|
|[<span data-ttu-id="d1469-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-910">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-911">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-911">Returns:</span></span>

<span data-ttu-id="d1469-p155">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d1469-914">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="d1469-914">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d1469-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d1469-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d1469-916">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-917">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d1469-p156">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d1469-921">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="d1469-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d1469-922">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="d1469-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d1469-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d1469-926">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-926">Requirements</span></span>

|<span data-ttu-id="d1469-927">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-927">Requirement</span></span>| <span data-ttu-id="d1469-928">値</span><span class="sxs-lookup"><span data-stu-id="d1469-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-929">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-930">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-930">1.0</span></span>|
|[<span data-ttu-id="d1469-931">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-932">ReadItem</span></span>|
|[<span data-ttu-id="d1469-933">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-934">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-935">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-935">Returns:</span></span>

<span data-ttu-id="d1469-p158">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="d1469-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="d1469-938">型: Object</span><span class="sxs-lookup"><span data-stu-id="d1469-938">Type: object</span></span>

##### <a name="example"></a><span data-ttu-id="d1469-939">例</span><span class="sxs-lookup"><span data-stu-id="d1469-939">Example</span></span>

<span data-ttu-id="d1469-940">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="d1469-940">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d1469-941">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="d1469-941">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d1469-942">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-942">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-943">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-943">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d1469-944">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="d1469-944">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d1469-p159">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="d1469-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-947">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-947">Parameters</span></span>

|<span data-ttu-id="d1469-948">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-948">Name</span></span>| <span data-ttu-id="d1469-949">種類</span><span class="sxs-lookup"><span data-stu-id="d1469-949">Type</span></span>| <span data-ttu-id="d1469-950">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-950">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="d1469-951">String</span><span class="sxs-lookup"><span data-stu-id="d1469-951">String</span></span>|<span data-ttu-id="d1469-952">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="d1469-952">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-953">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-953">Requirements</span></span>

|<span data-ttu-id="d1469-954">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-954">Requirement</span></span>| <span data-ttu-id="d1469-955">値</span><span class="sxs-lookup"><span data-stu-id="d1469-955">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-956">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-956">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-957">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-957">1.0</span></span>|
|[<span data-ttu-id="d1469-958">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-958">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-959">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-959">ReadItem</span></span>|
|[<span data-ttu-id="d1469-960">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-960">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-961">読み取り</span><span class="sxs-lookup"><span data-stu-id="d1469-961">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-962">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-962">Returns:</span></span>

<span data-ttu-id="d1469-963">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="d1469-963">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d1469-964">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d1469-964">Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d1469-965">例</span><span class="sxs-lookup"><span data-stu-id="d1469-965">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d1469-966">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d1469-966">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d1469-967">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-967">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d1469-p160">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-970">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-970">Parameters</span></span>

|<span data-ttu-id="d1469-971">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-971">Name</span></span>| <span data-ttu-id="d1469-972">型</span><span class="sxs-lookup"><span data-stu-id="d1469-972">Type</span></span>| <span data-ttu-id="d1469-973">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-973">Attributes</span></span>| <span data-ttu-id="d1469-974">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-974">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="d1469-975">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d1469-975">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d1469-p161">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="d1469-979">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-979">Object</span></span>| <span data-ttu-id="d1469-980">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-980">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-981">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-981">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d1469-982">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-982">Object</span></span>| <span data-ttu-id="d1469-983">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-983">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-984">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-984">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d1469-985">function</span><span class="sxs-lookup"><span data-stu-id="d1469-985">function</span></span>||<span data-ttu-id="d1469-986">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-986">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1469-987">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="d1469-987">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d1469-988">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="d1469-988">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-989">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-989">Requirements</span></span>

|<span data-ttu-id="d1469-990">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-990">Requirement</span></span>| <span data-ttu-id="d1469-991">値</span><span class="sxs-lookup"><span data-stu-id="d1469-991">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-992">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-992">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-993">1.2</span><span class="sxs-lookup"><span data-stu-id="d1469-993">1.2</span></span>|
|[<span data-ttu-id="d1469-994">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-994">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-995">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-995">ReadItem</span></span>|
|[<span data-ttu-id="d1469-996">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-996">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-997">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-997">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d1469-998">戻り値:</span><span class="sxs-lookup"><span data-stu-id="d1469-998">Returns:</span></span>

<span data-ttu-id="d1469-999">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="d1469-999">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d1469-1000">型:String</span><span class="sxs-lookup"><span data-stu-id="d1469-1000">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d1469-1001">例</span><span class="sxs-lookup"><span data-stu-id="d1469-1001">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d1469-1002">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d1469-1002">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d1469-1003">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1003">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d1469-p163">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="d1469-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-1007">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-1007">Parameters</span></span>

|<span data-ttu-id="d1469-1008">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-1008">Name</span></span>| <span data-ttu-id="d1469-1009">型</span><span class="sxs-lookup"><span data-stu-id="d1469-1009">Type</span></span>| <span data-ttu-id="d1469-1010">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-1010">Attributes</span></span>| <span data-ttu-id="d1469-1011">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-1011">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d1469-1012">function</span><span class="sxs-lookup"><span data-stu-id="d1469-1012">function</span></span>||<span data-ttu-id="d1469-1013">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1013">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1469-1014">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1014">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d1469-1015">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1015">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="d1469-1016">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1469-1016">Object</span></span>| <span data-ttu-id="d1469-1017">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1017">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1018">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1018">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d1469-1019">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1019">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-1020">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1020">Requirements</span></span>

|<span data-ttu-id="d1469-1021">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1021">Requirement</span></span>| <span data-ttu-id="d1469-1022">値</span><span class="sxs-lookup"><span data-stu-id="d1469-1022">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-1023">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-1023">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-1024">1.0</span><span class="sxs-lookup"><span data-stu-id="d1469-1024">1.0</span></span>|
|[<span data-ttu-id="d1469-1025">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-1025">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-1026">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d1469-1026">ReadItem</span></span>|
|[<span data-ttu-id="d1469-1027">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-1027">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-1028">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d1469-1028">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-1029">例</span><span class="sxs-lookup"><span data-stu-id="d1469-1029">Example</span></span>

<span data-ttu-id="d1469-p166">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="d1469-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d1469-1033">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d1469-1033">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d1469-1034">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1034">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d1469-1035">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1035">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d1469-1036">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="d1469-1036">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d1469-1037">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="d1469-1037">In Outlook on the web and OWA for Devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="d1469-1038">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1038">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-1039">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-1039">Parameters</span></span>

|<span data-ttu-id="d1469-1040">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-1040">Name</span></span>| <span data-ttu-id="d1469-1041">型</span><span class="sxs-lookup"><span data-stu-id="d1469-1041">Type</span></span>| <span data-ttu-id="d1469-1042">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-1042">Attributes</span></span>| <span data-ttu-id="d1469-1043">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-1043">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="d1469-1044">String</span><span class="sxs-lookup"><span data-stu-id="d1469-1044">String</span></span>||<span data-ttu-id="d1469-1045">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="d1469-1045">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="d1469-1046">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="d1469-1046">Object</span></span>| <span data-ttu-id="d1469-1047">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1047">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1048">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-1048">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d1469-1049">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-1049">Object</span></span>| <span data-ttu-id="d1469-1050">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1051">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1051">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d1469-1052">function</span><span class="sxs-lookup"><span data-stu-id="d1469-1052">function</span></span>| <span data-ttu-id="d1469-1053">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1054">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1054">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d1469-1055">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1055">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d1469-1056">エラー</span><span class="sxs-lookup"><span data-stu-id="d1469-1056">Errors</span></span>

| <span data-ttu-id="d1469-1057">エラー コード</span><span class="sxs-lookup"><span data-stu-id="d1469-1057">Error code</span></span> | <span data-ttu-id="d1469-1058">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-1058">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="d1469-1059">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="d1469-1059">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-1060">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1060">Requirements</span></span>

|<span data-ttu-id="d1469-1061">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1061">Requirement</span></span>| <span data-ttu-id="d1469-1062">値</span><span class="sxs-lookup"><span data-stu-id="d1469-1062">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-1063">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-1063">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-1064">1.1</span><span class="sxs-lookup"><span data-stu-id="d1469-1064">1.1</span></span>|
|[<span data-ttu-id="d1469-1065">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-1065">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-1066">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d1469-1066">ReadWriteItem</span></span>|
|[<span data-ttu-id="d1469-1067">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-1067">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-1068">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-1068">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-1069">例</span><span class="sxs-lookup"><span data-stu-id="d1469-1069">Example</span></span>

<span data-ttu-id="d1469-1070">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1070">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d1469-1071">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d1469-1071">saveAsync([options], callback)</span></span>

<span data-ttu-id="d1469-1072">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1072">Asynchronously saves an item.</span></span>

<span data-ttu-id="d1469-1073">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1073">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="d1469-1074">Outlook on the web またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1074">In Outlook Web App or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="d1469-1075">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1075">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-1076">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="d1469-1076">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d1469-1077">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1077">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d1469-p170">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d1469-1081">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="d1469-1081">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d1469-1082">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="d1469-1082">Note: Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="d1469-1083">`saveAsync` メソッドは、作成モードの会議から呼び出されると失敗します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1083">The `saveAsync` method will fail when called from a meeting in compose mode.</span></span> <span data-ttu-id="d1469-1084">回避策については、「[Office JS API を使用して Outlook for Mac で会議を下書きとして保存できない](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d1469-1084">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d1469-1085">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1085">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-1086">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-1086">Parameters</span></span>

|<span data-ttu-id="d1469-1087">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-1087">Name</span></span>| <span data-ttu-id="d1469-1088">型</span><span class="sxs-lookup"><span data-stu-id="d1469-1088">Type</span></span>| <span data-ttu-id="d1469-1089">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-1089">Attributes</span></span>| <span data-ttu-id="d1469-1090">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-1090">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="d1469-1091">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-1091">Object</span></span>| <span data-ttu-id="d1469-1092">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1093">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-1093">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d1469-1094">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-1094">Object</span></span>| <span data-ttu-id="d1469-1095">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1096">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1096">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="d1469-1097">function</span><span class="sxs-lookup"><span data-stu-id="d1469-1097">function</span></span>||<span data-ttu-id="d1469-1098">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1098">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d1469-1099">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1099">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d1469-1100">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1100">Requirements</span></span>

|<span data-ttu-id="d1469-1101">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1101">Requirement</span></span>| <span data-ttu-id="d1469-1102">値</span><span class="sxs-lookup"><span data-stu-id="d1469-1102">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-1103">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-1103">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-1104">1.3</span><span class="sxs-lookup"><span data-stu-id="d1469-1104">1.3</span></span>|
|[<span data-ttu-id="d1469-1105">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-1105">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-1106">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d1469-1106">ReadWriteItem</span></span>|
|[<span data-ttu-id="d1469-1107">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-1107">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-1108">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-1108">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d1469-1109">例</span><span class="sxs-lookup"><span data-stu-id="d1469-1109">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d1469-p172">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d1469-1112">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d1469-1112">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d1469-1113">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="d1469-1113">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d1469-p173">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d1469-1117">パラメーター</span><span class="sxs-lookup"><span data-stu-id="d1469-1117">Parameters</span></span>

|<span data-ttu-id="d1469-1118">名前</span><span class="sxs-lookup"><span data-stu-id="d1469-1118">Name</span></span>| <span data-ttu-id="d1469-1119">型</span><span class="sxs-lookup"><span data-stu-id="d1469-1119">Type</span></span>| <span data-ttu-id="d1469-1120">属性</span><span class="sxs-lookup"><span data-stu-id="d1469-1120">Attributes</span></span>| <span data-ttu-id="d1469-1121">説明</span><span class="sxs-lookup"><span data-stu-id="d1469-1121">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d1469-1122">String</span><span class="sxs-lookup"><span data-stu-id="d1469-1122">String</span></span>||<span data-ttu-id="d1469-p174">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="d1469-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="d1469-1126">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-1126">Object</span></span>| <span data-ttu-id="d1469-1127">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1127">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1128">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="d1469-1128">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="d1469-1129">Object</span><span class="sxs-lookup"><span data-stu-id="d1469-1129">Object</span></span>| <span data-ttu-id="d1469-1130">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1131">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1131">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d1469-1132">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d1469-1132">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d1469-1133">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d1469-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="d1469-1134">`text` の場合、Outlook on the web とデスクトップ クライアントでは現在のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1134">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d1469-1135">フィールドが HTML エディターの場合、データが HTML の場合でもテキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1135">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d1469-1136">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Outlook on the web では現在のスタイルが適用され、Outlook デスクトップ クライアントでは既定のスタイルが適用されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1136">If `html` and the field supports HTML (the subject doesn&#39;t), the current style is applied in Outlook Web App and the default style is applied in Outlook.</span></span> <span data-ttu-id="d1469-1137">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1137">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d1469-1138">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1138">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="d1469-1139">function</span><span class="sxs-lookup"><span data-stu-id="d1469-1139">function</span></span>||<span data-ttu-id="d1469-1140">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="d1469-1140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d1469-1141">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1141">Requirements</span></span>

|<span data-ttu-id="d1469-1142">要件</span><span class="sxs-lookup"><span data-stu-id="d1469-1142">Requirement</span></span>| <span data-ttu-id="d1469-1143">値</span><span class="sxs-lookup"><span data-stu-id="d1469-1143">Value</span></span>|
|---|---|
|[<span data-ttu-id="d1469-1144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d1469-1144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d1469-1145">1.2</span><span class="sxs-lookup"><span data-stu-id="d1469-1145">1.2</span></span>|
|[<span data-ttu-id="d1469-1146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d1469-1146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d1469-1147">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d1469-1147">ReadWriteItem</span></span>|
|[<span data-ttu-id="d1469-1148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d1469-1148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d1469-1149">作成</span><span class="sxs-lookup"><span data-stu-id="d1469-1149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d1469-1150">例</span><span class="sxs-lookup"><span data-stu-id="d1469-1150">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
