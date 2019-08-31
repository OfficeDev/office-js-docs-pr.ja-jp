---
title: Office. メールボックス-要件セット1.1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20d3aaecc5e0c62f86a46ae29010a6462446bf1d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696443"
---
# <a name="item"></a><span data-ttu-id="bfb6a-102">item</span><span class="sxs-lookup"><span data-stu-id="bfb6a-102">item</span></span>

### <span data-ttu-id="bfb6a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="bfb6a-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-107">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-107">Requirements</span></span>

|<span data-ttu-id="bfb6a-108">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-108">Requirement</span></span>| <span data-ttu-id="bfb6a-109">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-111">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-111">1.0</span></span>|
|[<span data-ttu-id="bfb6a-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="bfb6a-113">Restricted</span></span>|
|[<span data-ttu-id="bfb6a-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bfb6a-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-116">Members and methods</span></span>

| <span data-ttu-id="bfb6a-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-117">Member</span></span> | <span data-ttu-id="bfb6a-118">種類</span><span class="sxs-lookup"><span data-stu-id="bfb6a-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bfb6a-119">attachments</span><span class="sxs-lookup"><span data-stu-id="bfb6a-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="bfb6a-120">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-120">Member</span></span> |
| [<span data-ttu-id="bfb6a-121">bcc</span><span class="sxs-lookup"><span data-stu-id="bfb6a-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="bfb6a-122">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-122">Member</span></span> |
| [<span data-ttu-id="bfb6a-123">body</span><span class="sxs-lookup"><span data-stu-id="bfb6a-123">body</span></span>](#body-body) | <span data-ttu-id="bfb6a-124">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-124">Member</span></span> |
| [<span data-ttu-id="bfb6a-125">cc</span><span class="sxs-lookup"><span data-stu-id="bfb6a-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bfb6a-126">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-126">Member</span></span> |
| [<span data-ttu-id="bfb6a-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="bfb6a-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="bfb6a-128">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-128">Member</span></span> |
| [<span data-ttu-id="bfb6a-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="bfb6a-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="bfb6a-130">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-130">Member</span></span> |
| [<span data-ttu-id="bfb6a-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="bfb6a-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="bfb6a-132">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-132">Member</span></span> |
| [<span data-ttu-id="bfb6a-133">end</span><span class="sxs-lookup"><span data-stu-id="bfb6a-133">end</span></span>](#end-datetime) | <span data-ttu-id="bfb6a-134">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-134">Member</span></span> |
| [<span data-ttu-id="bfb6a-135">from</span><span class="sxs-lookup"><span data-stu-id="bfb6a-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="bfb6a-136">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-136">Member</span></span> |
| [<span data-ttu-id="bfb6a-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="bfb6a-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="bfb6a-138">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-138">Member</span></span> |
| [<span data-ttu-id="bfb6a-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="bfb6a-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="bfb6a-140">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-140">Member</span></span> |
| [<span data-ttu-id="bfb6a-141">itemId</span><span class="sxs-lookup"><span data-stu-id="bfb6a-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="bfb6a-142">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-142">Member</span></span> |
| [<span data-ttu-id="bfb6a-143">itemType</span><span class="sxs-lookup"><span data-stu-id="bfb6a-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="bfb6a-144">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-144">Member</span></span> |
| [<span data-ttu-id="bfb6a-145">location</span><span class="sxs-lookup"><span data-stu-id="bfb6a-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="bfb6a-146">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-146">Member</span></span> |
| [<span data-ttu-id="bfb6a-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="bfb6a-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="bfb6a-148">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-148">Member</span></span> |
| [<span data-ttu-id="bfb6a-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="bfb6a-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bfb6a-150">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-150">Member</span></span> |
| [<span data-ttu-id="bfb6a-151">organizer</span><span class="sxs-lookup"><span data-stu-id="bfb6a-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="bfb6a-152">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-152">Member</span></span> |
| [<span data-ttu-id="bfb6a-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="bfb6a-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bfb6a-154">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-154">Member</span></span> |
| [<span data-ttu-id="bfb6a-155">sender</span><span class="sxs-lookup"><span data-stu-id="bfb6a-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="bfb6a-156">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-156">Member</span></span> |
| [<span data-ttu-id="bfb6a-157">start</span><span class="sxs-lookup"><span data-stu-id="bfb6a-157">start</span></span>](#start-datetime) | <span data-ttu-id="bfb6a-158">Member</span><span class="sxs-lookup"><span data-stu-id="bfb6a-158">Member</span></span> |
| [<span data-ttu-id="bfb6a-159">subject</span><span class="sxs-lookup"><span data-stu-id="bfb6a-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="bfb6a-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-160">Member</span></span> |
| [<span data-ttu-id="bfb6a-161">to</span><span class="sxs-lookup"><span data-stu-id="bfb6a-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="bfb6a-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-162">Member</span></span> |
| [<span data-ttu-id="bfb6a-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bfb6a-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="bfb6a-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-164">Method</span></span> |
| [<span data-ttu-id="bfb6a-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bfb6a-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="bfb6a-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-166">Method</span></span> |
| [<span data-ttu-id="bfb6a-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="bfb6a-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="bfb6a-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-168">Method</span></span> |
| [<span data-ttu-id="bfb6a-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="bfb6a-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="bfb6a-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-170">Method</span></span> |
| [<span data-ttu-id="bfb6a-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="bfb6a-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="bfb6a-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-172">Method</span></span> |
| [<span data-ttu-id="bfb6a-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="bfb6a-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bfb6a-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-174">Method</span></span> |
| [<span data-ttu-id="bfb6a-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="bfb6a-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="bfb6a-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-176">Method</span></span> |
| [<span data-ttu-id="bfb6a-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="bfb6a-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="bfb6a-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-178">Method</span></span> |
| [<span data-ttu-id="bfb6a-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="bfb6a-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="bfb6a-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-180">Method</span></span> |
| [<span data-ttu-id="bfb6a-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="bfb6a-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="bfb6a-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-182">Method</span></span> |
| [<span data-ttu-id="bfb6a-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="bfb6a-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="bfb6a-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="bfb6a-185">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-185">Example</span></span>

<span data-ttu-id="bfb6a-186">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="bfb6a-187">メンバー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-188">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="bfb6a-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="bfb6a-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-191">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="bfb6a-192">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-193">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-193">Type</span></span>

*   <span data-ttu-id="bfb6a-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="bfb6a-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-195">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-195">Requirements</span></span>

|<span data-ttu-id="bfb6a-196">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-196">Requirement</span></span>| <span data-ttu-id="bfb6a-197">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-199">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-199">1.0</span></span>|
|[<span data-ttu-id="bfb6a-200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-201">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-203">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-204">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-204">Example</span></span>

<span data-ttu-id="bfb6a-205">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-206">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-207">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="bfb6a-208">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-209">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-209">Type</span></span>

*   [<span data-ttu-id="bfb6a-210">受信者</span><span class="sxs-lookup"><span data-stu-id="bfb6a-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-211">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-211">Requirements</span></span>

|<span data-ttu-id="bfb6a-212">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-212">Requirement</span></span>| <span data-ttu-id="bfb6a-213">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-215">1.1</span><span class="sxs-lookup"><span data-stu-id="bfb6a-215">1.1</span></span>|
|[<span data-ttu-id="bfb6a-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-217">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-219">作成</span><span class="sxs-lookup"><span data-stu-id="bfb6a-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-220">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-220">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="bfb6a-221">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-222">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-223">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-223">Type</span></span>

*   [<span data-ttu-id="bfb6a-224">Body</span><span class="sxs-lookup"><span data-stu-id="bfb6a-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-225">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-225">Requirements</span></span>

|<span data-ttu-id="bfb6a-226">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-226">Requirement</span></span>| <span data-ttu-id="bfb6a-227">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-228">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-229">1.1</span><span class="sxs-lookup"><span data-stu-id="bfb6a-229">1.1</span></span>|
|[<span data-ttu-id="bfb6a-230">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-231">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-232">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-233">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-234">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-234">Example</span></span>

<span data-ttu-id="bfb6a-235">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-235">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="bfb6a-236">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-236">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-237">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-238">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="bfb6a-239">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-240">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-240">Read mode</span></span>

<span data-ttu-id="bfb6a-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-243">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-243">Compose mode</span></span>

<span data-ttu-id="bfb6a-244">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-245">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-245">Type</span></span>

*   <span data-ttu-id="bfb6a-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-247">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-247">Requirements</span></span>

|<span data-ttu-id="bfb6a-248">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-248">Requirement</span></span>| <span data-ttu-id="bfb6a-249">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-251">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-251">1.0</span></span>|
|[<span data-ttu-id="bfb6a-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-253">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="bfb6a-256">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="bfb6a-257">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="bfb6a-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="bfb6a-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-262">Type</span><span class="sxs-lookup"><span data-stu-id="bfb6a-262">Type</span></span>

*   <span data-ttu-id="bfb6a-263">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-264">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-264">Requirements</span></span>

|<span data-ttu-id="bfb6a-265">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-265">Requirement</span></span>| <span data-ttu-id="bfb6a-266">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-268">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-268">1.0</span></span>|
|[<span data-ttu-id="bfb6a-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-270">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-273">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-273">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="bfb6a-274">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="bfb6a-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="bfb6a-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-277">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-277">Type</span></span>

*   <span data-ttu-id="bfb6a-278">日付</span><span class="sxs-lookup"><span data-stu-id="bfb6a-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-279">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-279">Requirements</span></span>

|<span data-ttu-id="bfb6a-280">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-280">Requirement</span></span>| <span data-ttu-id="bfb6a-281">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-283">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-283">1.0</span></span>|
|[<span data-ttu-id="bfb6a-284">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-285">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-287">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-288">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-288">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="bfb6a-289">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="bfb6a-289">dateTimeModified: Date</span></span>

<span data-ttu-id="bfb6a-290">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="bfb6a-291">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="bfb6a-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-292">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-293">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-293">Type</span></span>

*   <span data-ttu-id="bfb6a-294">日付</span><span class="sxs-lookup"><span data-stu-id="bfb6a-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-295">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-295">Requirements</span></span>

|<span data-ttu-id="bfb6a-296">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-296">Requirement</span></span>| <span data-ttu-id="bfb6a-297">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-299">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-299">1.0</span></span>|
|[<span data-ttu-id="bfb6a-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-301">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-303">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-304">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-304">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="bfb6a-305">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-306">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="bfb6a-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-309">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-309">Read mode</span></span>

<span data-ttu-id="bfb6a-310">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-310">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-311">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-311">Compose mode</span></span>

<span data-ttu-id="bfb6a-312">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="bfb6a-313">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bfb6a-314">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="bfb6a-315">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-315">Type</span></span>

*   <span data-ttu-id="bfb6a-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-317">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-317">Requirements</span></span>

|<span data-ttu-id="bfb6a-318">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-318">Requirement</span></span>| <span data-ttu-id="bfb6a-319">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-320">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-321">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-321">1.0</span></span>|
|[<span data-ttu-id="bfb6a-322">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-323">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-324">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-325">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-325">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-326">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="bfb6a-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-331">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-332">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-332">Type</span></span>

*   [<span data-ttu-id="bfb6a-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bfb6a-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-334">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-334">Requirements</span></span>

|<span data-ttu-id="bfb6a-335">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-335">Requirement</span></span>| <span data-ttu-id="bfb6a-336">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-337">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-338">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-338">1.0</span></span>|
|[<span data-ttu-id="bfb6a-339">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-340">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-341">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-342">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-343">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-343">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="bfb6a-344">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-344">internetMessageId: String</span></span>

<span data-ttu-id="bfb6a-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-347">Type</span><span class="sxs-lookup"><span data-stu-id="bfb6a-347">Type</span></span>

*   <span data-ttu-id="bfb6a-348">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-349">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-349">Requirements</span></span>

|<span data-ttu-id="bfb6a-350">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-350">Requirement</span></span>| <span data-ttu-id="bfb6a-351">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-352">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-353">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-353">1.0</span></span>|
|[<span data-ttu-id="bfb6a-354">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-355">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-356">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-357">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-358">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-358">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="bfb6a-359">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-359">itemClass: String</span></span>

<span data-ttu-id="bfb6a-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="bfb6a-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="bfb6a-364">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-364">Type</span></span> | <span data-ttu-id="bfb6a-365">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-365">Description</span></span> | <span data-ttu-id="bfb6a-366">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="bfb6a-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="bfb6a-367">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="bfb6a-367">Appointment items</span></span> | <span data-ttu-id="bfb6a-368">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="bfb6a-369">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="bfb6a-369">Message items</span></span> | <span data-ttu-id="bfb6a-370">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="bfb6a-371">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-372">Type</span><span class="sxs-lookup"><span data-stu-id="bfb6a-372">Type</span></span>

*   <span data-ttu-id="bfb6a-373">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-374">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-374">Requirements</span></span>

|<span data-ttu-id="bfb6a-375">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-375">Requirement</span></span>| <span data-ttu-id="bfb6a-376">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-377">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-378">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-378">1.0</span></span>|
|[<span data-ttu-id="bfb6a-379">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-380">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-381">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-382">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-383">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-383">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="bfb6a-384">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-384">(nullable) itemId: String</span></span>

<span data-ttu-id="bfb6a-385">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="bfb6a-386">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-387">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="bfb6a-388">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="bfb6a-389">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="bfb6a-390">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-391">Type</span><span class="sxs-lookup"><span data-stu-id="bfb6a-391">Type</span></span>

*   <span data-ttu-id="bfb6a-392">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-393">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-393">Requirements</span></span>

|<span data-ttu-id="bfb6a-394">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-394">Requirement</span></span>| <span data-ttu-id="bfb6a-395">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-396">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-397">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-397">1.0</span></span>|
|[<span data-ttu-id="bfb6a-398">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-399">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-400">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-401">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-402">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-402">Example</span></span>

<span data-ttu-id="bfb6a-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="bfb6a-405">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-406">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="bfb6a-407">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-408">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-408">Type</span></span>

*   [<span data-ttu-id="bfb6a-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="bfb6a-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-410">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-410">Requirements</span></span>

|<span data-ttu-id="bfb6a-411">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-411">Requirement</span></span>| <span data-ttu-id="bfb6a-412">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-413">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-414">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-414">1.0</span></span>|
|[<span data-ttu-id="bfb6a-415">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-416">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-417">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-418">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-419">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-419">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="bfb6a-420">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-421">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-422">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-422">Read mode</span></span>

<span data-ttu-id="bfb6a-423">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-424">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-424">Compose mode</span></span>

<span data-ttu-id="bfb6a-425">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-426">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-426">Type</span></span>

*   <span data-ttu-id="bfb6a-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-428">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-428">Requirements</span></span>

|<span data-ttu-id="bfb6a-429">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-429">Requirement</span></span>| <span data-ttu-id="bfb6a-430">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-431">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-432">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-432">1.0</span></span>|
|[<span data-ttu-id="bfb6a-433">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-434">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-435">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-436">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-436">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="bfb6a-437">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-437">normalizedSubject: String</span></span>

<span data-ttu-id="bfb6a-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="bfb6a-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-442">Type</span><span class="sxs-lookup"><span data-stu-id="bfb6a-442">Type</span></span>

*   <span data-ttu-id="bfb6a-443">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-444">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-444">Requirements</span></span>

|<span data-ttu-id="bfb6a-445">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-445">Requirement</span></span>| <span data-ttu-id="bfb6a-446">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-448">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-448">1.0</span></span>|
|[<span data-ttu-id="bfb6a-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-450">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-453">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-453">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-454">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-455">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="bfb6a-456">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-457">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-457">Read mode</span></span>

<span data-ttu-id="bfb6a-458">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-459">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-459">Compose mode</span></span>

<span data-ttu-id="bfb6a-460">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-461">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-461">Type</span></span>

*   <span data-ttu-id="bfb6a-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-463">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-463">Requirements</span></span>

|<span data-ttu-id="bfb6a-464">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-464">Requirement</span></span>| <span data-ttu-id="bfb6a-465">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-466">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-467">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-467">1.0</span></span>|
|[<span data-ttu-id="bfb6a-468">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-469">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-470">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-471">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-471">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-472">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-475">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-475">Type</span></span>

*   [<span data-ttu-id="bfb6a-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bfb6a-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-477">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-477">Requirements</span></span>

|<span data-ttu-id="bfb6a-478">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-478">Requirement</span></span>| <span data-ttu-id="bfb6a-479">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-480">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-481">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-481">1.0</span></span>|
|[<span data-ttu-id="bfb6a-482">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-483">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-484">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-485">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-486">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-486">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-487">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-488">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="bfb6a-489">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-490">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-490">Read mode</span></span>

<span data-ttu-id="bfb6a-491">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-492">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-492">Compose mode</span></span>

<span data-ttu-id="bfb6a-493">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-494">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-494">Type</span></span>

*   <span data-ttu-id="bfb6a-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-496">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-496">Requirements</span></span>

|<span data-ttu-id="bfb6a-497">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-497">Requirement</span></span>| <span data-ttu-id="bfb6a-498">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-499">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-500">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-500">1.0</span></span>|
|[<span data-ttu-id="bfb6a-501">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-502">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-503">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-504">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-504">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-505">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="bfb6a-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-510">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="bfb6a-511">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-511">Type</span></span>

*   [<span data-ttu-id="bfb6a-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="bfb6a-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="bfb6a-513">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-513">Requirements</span></span>

|<span data-ttu-id="bfb6a-514">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-514">Requirement</span></span>| <span data-ttu-id="bfb6a-515">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-516">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-517">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-517">1.0</span></span>|
|[<span data-ttu-id="bfb6a-518">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-519">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-520">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-521">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-522">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-522">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="bfb6a-523">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-524">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="bfb6a-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-527">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-527">Read mode</span></span>

<span data-ttu-id="bfb6a-528">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-528">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-529">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-529">Compose mode</span></span>

<span data-ttu-id="bfb6a-530">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="bfb6a-531">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="bfb6a-532">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="bfb6a-533">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-533">Type</span></span>

*   <span data-ttu-id="bfb6a-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-535">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-535">Requirements</span></span>

|<span data-ttu-id="bfb6a-536">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-536">Requirement</span></span>| <span data-ttu-id="bfb6a-537">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-538">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-539">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-539">1.0</span></span>|
|[<span data-ttu-id="bfb6a-540">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-541">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-542">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-543">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-543">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="bfb6a-544">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-545">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="bfb6a-546">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-547">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-547">Read mode</span></span>

<span data-ttu-id="bfb6a-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-550">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-550">Compose mode</span></span>

<span data-ttu-id="bfb6a-551">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-552">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-552">Type</span></span>

*   <span data-ttu-id="bfb6a-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-554">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-554">Requirements</span></span>

|<span data-ttu-id="bfb6a-555">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-555">Requirement</span></span>| <span data-ttu-id="bfb6a-556">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-557">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-558">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-558">1.0</span></span>|
|[<span data-ttu-id="bfb6a-559">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-560">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-561">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-562">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-562">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="bfb6a-563">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="bfb6a-564">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="bfb6a-565">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="bfb6a-566">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-566">Read mode</span></span>

<span data-ttu-id="bfb6a-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="bfb6a-569">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-569">Compose mode</span></span>

<span data-ttu-id="bfb6a-570">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="bfb6a-571">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-571">Type</span></span>

*   <span data-ttu-id="bfb6a-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-573">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-573">Requirements</span></span>

|<span data-ttu-id="bfb6a-574">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-574">Requirement</span></span>| <span data-ttu-id="bfb6a-575">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-576">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-577">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-577">1.0</span></span>|
|[<span data-ttu-id="bfb6a-578">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-579">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-580">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-581">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="bfb6a-582">メソッド</span><span class="sxs-lookup"><span data-stu-id="bfb6a-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="bfb6a-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bfb6a-584">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="bfb6a-585">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="bfb6a-586">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-587">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-587">Parameters</span></span>

|<span data-ttu-id="bfb6a-588">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-588">Name</span></span>| <span data-ttu-id="bfb6a-589">種類</span><span class="sxs-lookup"><span data-stu-id="bfb6a-589">Type</span></span>| <span data-ttu-id="bfb6a-590">属性</span><span class="sxs-lookup"><span data-stu-id="bfb6a-590">Attributes</span></span>| <span data-ttu-id="bfb6a-591">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="bfb6a-592">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-592">String</span></span>||<span data-ttu-id="bfb6a-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bfb6a-595">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-595">String</span></span>||<span data-ttu-id="bfb6a-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bfb6a-598">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-598">Object</span></span>| <span data-ttu-id="bfb6a-599">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-599">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-600">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bfb6a-601">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-601">Object</span></span>| <span data-ttu-id="bfb6a-602">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-602">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-603">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bfb6a-604">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-604">function</span></span>| <span data-ttu-id="bfb6a-605">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-605">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-606">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bfb6a-607">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bfb6a-608">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bfb6a-609">エラー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-609">Errors</span></span>

| <span data-ttu-id="bfb6a-610">エラー コード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-610">Error code</span></span> | <span data-ttu-id="bfb6a-611">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="bfb6a-612">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="bfb6a-613">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bfb6a-614">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bfb6a-615">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-615">Requirements</span></span>

|<span data-ttu-id="bfb6a-616">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-616">Requirement</span></span>| <span data-ttu-id="bfb6a-617">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-618">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-619">1.1</span><span class="sxs-lookup"><span data-stu-id="bfb6a-619">1.1</span></span>|
|[<span data-ttu-id="bfb6a-620">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="bfb6a-622">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-623">作成</span><span class="sxs-lookup"><span data-stu-id="bfb6a-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-624">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-624">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="bfb6a-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="bfb6a-626">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="bfb6a-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="bfb6a-630">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="bfb6a-631">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-632">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-632">Parameters</span></span>

|<span data-ttu-id="bfb6a-633">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-633">Name</span></span>| <span data-ttu-id="bfb6a-634">種類</span><span class="sxs-lookup"><span data-stu-id="bfb6a-634">Type</span></span>| <span data-ttu-id="bfb6a-635">属性</span><span class="sxs-lookup"><span data-stu-id="bfb6a-635">Attributes</span></span>| <span data-ttu-id="bfb6a-636">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="bfb6a-637">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-637">String</span></span>||<span data-ttu-id="bfb6a-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="bfb6a-640">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-640">String</span></span>||<span data-ttu-id="bfb6a-641">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-641">The subject of the item to be attached.</span></span> <span data-ttu-id="bfb6a-642">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="bfb6a-643">Object</span><span class="sxs-lookup"><span data-stu-id="bfb6a-643">Object</span></span>| <span data-ttu-id="bfb6a-644">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-644">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-645">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bfb6a-646">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-646">Object</span></span>| <span data-ttu-id="bfb6a-647">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-647">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-648">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bfb6a-649">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-649">function</span></span>| <span data-ttu-id="bfb6a-650">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-650">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-651">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bfb6a-652">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="bfb6a-653">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bfb6a-654">エラー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-654">Errors</span></span>

| <span data-ttu-id="bfb6a-655">エラー コード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-655">Error code</span></span> | <span data-ttu-id="bfb6a-656">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="bfb6a-657">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bfb6a-658">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-658">Requirements</span></span>

|<span data-ttu-id="bfb6a-659">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-659">Requirement</span></span>| <span data-ttu-id="bfb6a-660">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-661">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-662">1.1</span><span class="sxs-lookup"><span data-stu-id="bfb6a-662">1.1</span></span>|
|[<span data-ttu-id="bfb6a-663">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="bfb6a-665">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-666">作成</span><span class="sxs-lookup"><span data-stu-id="bfb6a-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-667">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-667">Example</span></span>

<span data-ttu-id="bfb6a-668">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="bfb6a-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="bfb6a-670">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-671">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bfb6a-672">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bfb6a-673">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-674">へ`displayReplyAllForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="bfb6a-675">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-676">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-676">Parameters</span></span>

|<span data-ttu-id="bfb6a-677">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-677">Name</span></span>| <span data-ttu-id="bfb6a-678">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-678">Type</span></span>| <span data-ttu-id="bfb6a-679">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="bfb6a-680">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bfb6a-680">String &#124; Object</span></span>| |<span data-ttu-id="bfb6a-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bfb6a-683">**または**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-683">**OR**</span></span><br/><span data-ttu-id="bfb6a-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bfb6a-686">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-686">String</span></span> | <span data-ttu-id="bfb6a-687">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-687">&lt;optional&gt;</span></span> | <span data-ttu-id="bfb6a-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="bfb6a-690">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-690">function</span></span> | <span data-ttu-id="bfb6a-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-691">&lt;optional&gt;</span></span> | <span data-ttu-id="bfb6a-692">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bfb6a-693">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-693">Requirements</span></span>

|<span data-ttu-id="bfb6a-694">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-694">Requirement</span></span>| <span data-ttu-id="bfb6a-695">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-696">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-697">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-697">1.0</span></span>|
|[<span data-ttu-id="bfb6a-698">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-699">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-700">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-701">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bfb6a-702">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-702">Examples</span></span>

<span data-ttu-id="bfb6a-703">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="bfb6a-704">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-704">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="bfb6a-705">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-705">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bfb6a-706">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-706">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="bfb6a-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="bfb6a-708">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-709">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bfb6a-710">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="bfb6a-711">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-712">へ`displayReplyForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="bfb6a-713">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-714">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-714">Parameters</span></span>

|<span data-ttu-id="bfb6a-715">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-715">Name</span></span>| <span data-ttu-id="bfb6a-716">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-716">Type</span></span>| <span data-ttu-id="bfb6a-717">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="bfb6a-718">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="bfb6a-718">String &#124; Object</span></span>| | <span data-ttu-id="bfb6a-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="bfb6a-721">**または**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-721">**OR**</span></span><br/><span data-ttu-id="bfb6a-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="bfb6a-724">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-724">String</span></span> | <span data-ttu-id="bfb6a-725">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-725">&lt;optional&gt;</span></span> | <span data-ttu-id="bfb6a-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="bfb6a-728">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-728">function</span></span> | <span data-ttu-id="bfb6a-729">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-729">&lt;optional&gt;</span></span> | <span data-ttu-id="bfb6a-730">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bfb6a-731">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-731">Requirements</span></span>

|<span data-ttu-id="bfb6a-732">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-732">Requirement</span></span>| <span data-ttu-id="bfb6a-733">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-734">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-735">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-735">1.0</span></span>|
|[<span data-ttu-id="bfb6a-736">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-737">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-738">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-739">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="bfb6a-740">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-740">Examples</span></span>

<span data-ttu-id="bfb6a-741">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="bfb6a-742">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-742">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="bfb6a-743">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-743">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="bfb6a-744">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-744">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="bfb6a-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="bfb6a-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="bfb6a-746">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-747">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-748">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-748">Requirements</span></span>

|<span data-ttu-id="bfb6a-749">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-749">Requirement</span></span>| <span data-ttu-id="bfb6a-750">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-751">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-752">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-752">1.0</span></span>|
|[<span data-ttu-id="bfb6a-753">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-754">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-755">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-756">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bfb6a-757">戻り値:</span><span class="sxs-lookup"><span data-stu-id="bfb6a-757">Returns:</span></span>

<span data-ttu-id="bfb6a-758">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="bfb6a-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="bfb6a-759">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-759">Example</span></span>

<span data-ttu-id="bfb6a-760">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-760">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="bfb6a-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="bfb6a-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="bfb6a-762">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-763">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-764">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-764">Parameters</span></span>

|<span data-ttu-id="bfb6a-765">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-765">Name</span></span>| <span data-ttu-id="bfb6a-766">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-766">Type</span></span>| <span data-ttu-id="bfb6a-767">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="bfb6a-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="bfb6a-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="bfb6a-769">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bfb6a-770">Requirements</span><span class="sxs-lookup"><span data-stu-id="bfb6a-770">Requirements</span></span>

|<span data-ttu-id="bfb6a-771">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-771">Requirement</span></span>| <span data-ttu-id="bfb6a-772">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-773">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-774">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-774">1.0</span></span>|
|[<span data-ttu-id="bfb6a-775">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-776">制限あり</span><span class="sxs-lookup"><span data-stu-id="bfb6a-776">Restricted</span></span>|
|[<span data-ttu-id="bfb6a-777">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-778">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bfb6a-779">戻り値:</span><span class="sxs-lookup"><span data-stu-id="bfb6a-779">Returns:</span></span>

<span data-ttu-id="bfb6a-780">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="bfb6a-781">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="bfb6a-782">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="bfb6a-783">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="bfb6a-784">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-784">Value of `entityType`</span></span> | <span data-ttu-id="bfb6a-785">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-785">Type of objects in returned array</span></span> | <span data-ttu-id="bfb6a-786">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="bfb6a-787">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-787">String</span></span> | <span data-ttu-id="bfb6a-788">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="bfb6a-789">連絡先</span><span class="sxs-lookup"><span data-stu-id="bfb6a-789">Contact</span></span> | <span data-ttu-id="bfb6a-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="bfb6a-791">文字列</span><span class="sxs-lookup"><span data-stu-id="bfb6a-791">String</span></span> | <span data-ttu-id="bfb6a-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="bfb6a-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="bfb6a-793">MeetingSuggestion</span></span> | <span data-ttu-id="bfb6a-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="bfb6a-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="bfb6a-795">PhoneNumber</span></span> | <span data-ttu-id="bfb6a-796">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="bfb6a-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="bfb6a-797">TaskSuggestion</span></span> | <span data-ttu-id="bfb6a-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="bfb6a-799">文字列</span><span class="sxs-lookup"><span data-stu-id="bfb6a-799">String</span></span> | <span data-ttu-id="bfb6a-800">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="bfb6a-800">**Restricted**</span></span> |

<span data-ttu-id="bfb6a-801">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="bfb6a-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="bfb6a-802">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-802">Example</span></span>

<span data-ttu-id="bfb6a-803">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="bfb6a-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="bfb6a-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="bfb6a-805">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-806">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bfb6a-807">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-808">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-808">Parameters</span></span>

|<span data-ttu-id="bfb6a-809">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-809">Name</span></span>| <span data-ttu-id="bfb6a-810">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-810">Type</span></span>| <span data-ttu-id="bfb6a-811">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bfb6a-812">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-812">String</span></span>|<span data-ttu-id="bfb6a-813">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bfb6a-814">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-814">Requirements</span></span>

|<span data-ttu-id="bfb6a-815">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-815">Requirement</span></span>| <span data-ttu-id="bfb6a-816">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-817">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-818">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-818">1.0</span></span>|
|[<span data-ttu-id="bfb6a-819">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-820">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-821">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-822">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bfb6a-823">戻り値:</span><span class="sxs-lookup"><span data-stu-id="bfb6a-823">Returns:</span></span>

<span data-ttu-id="bfb6a-p146">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="bfb6a-826">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="bfb6a-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="bfb6a-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="bfb6a-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="bfb6a-828">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-829">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bfb6a-p147">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="bfb6a-833">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="bfb6a-834">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="bfb6a-p148">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bfb6a-837">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-837">Requirements</span></span>

|<span data-ttu-id="bfb6a-838">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-838">Requirement</span></span>| <span data-ttu-id="bfb6a-839">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-840">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-841">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-841">1.0</span></span>|
|[<span data-ttu-id="bfb6a-842">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-843">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-844">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-845">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bfb6a-846">戻り値:</span><span class="sxs-lookup"><span data-stu-id="bfb6a-846">Returns:</span></span>

<span data-ttu-id="bfb6a-p149">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="bfb6a-849">型: Object</span><span class="sxs-lookup"><span data-stu-id="bfb6a-849">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="bfb6a-850">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-850">Example</span></span>

<span data-ttu-id="bfb6a-851">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="bfb6a-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="bfb6a-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="bfb6a-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="bfb6a-853">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="bfb6a-854">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-854">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bfb6a-855">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="bfb6a-p150">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-858">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-858">Parameters</span></span>

|<span data-ttu-id="bfb6a-859">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-859">Name</span></span>| <span data-ttu-id="bfb6a-860">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-860">Type</span></span>| <span data-ttu-id="bfb6a-861">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="bfb6a-862">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-862">String</span></span>|<span data-ttu-id="bfb6a-863">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bfb6a-864">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-864">Requirements</span></span>

|<span data-ttu-id="bfb6a-865">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-865">Requirement</span></span>| <span data-ttu-id="bfb6a-866">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-867">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-868">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-868">1.0</span></span>|
|[<span data-ttu-id="bfb6a-869">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-869">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-870">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-871">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-871">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-872">読み取り</span><span class="sxs-lookup"><span data-stu-id="bfb6a-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bfb6a-873">戻り値:</span><span class="sxs-lookup"><span data-stu-id="bfb6a-873">Returns:</span></span>

<span data-ttu-id="bfb6a-874">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="bfb6a-875">型: Array. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="bfb6a-875">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="bfb6a-876">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="bfb6a-877">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-877">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="bfb6a-878">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-878">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="bfb6a-p151">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-882">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-882">Parameters</span></span>

|<span data-ttu-id="bfb6a-883">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-883">Name</span></span>| <span data-ttu-id="bfb6a-884">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-884">Type</span></span>| <span data-ttu-id="bfb6a-885">属性</span><span class="sxs-lookup"><span data-stu-id="bfb6a-885">Attributes</span></span>| <span data-ttu-id="bfb6a-886">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-886">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bfb6a-887">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-887">function</span></span>||<span data-ttu-id="bfb6a-888">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-888">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bfb6a-889">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-889">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="bfb6a-890">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-890">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="bfb6a-891">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-891">Object</span></span>| <span data-ttu-id="bfb6a-892">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-892">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-893">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-893">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="bfb6a-894">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-894">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bfb6a-895">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-895">Requirements</span></span>

|<span data-ttu-id="bfb6a-896">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-896">Requirement</span></span>| <span data-ttu-id="bfb6a-897">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-897">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-898">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-898">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-899">1.0</span><span class="sxs-lookup"><span data-stu-id="bfb6a-899">1.0</span></span>|
|[<span data-ttu-id="bfb6a-900">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-900">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-901">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-901">ReadItem</span></span>|
|[<span data-ttu-id="bfb6a-902">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-902">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-903">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bfb6a-903">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-904">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-904">Example</span></span>

<span data-ttu-id="bfb6a-p154">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="bfb6a-908">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="bfb6a-908">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="bfb6a-909">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-909">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="bfb6a-910">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-910">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="bfb6a-911">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-911">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="bfb6a-912">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-912">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="bfb6a-913">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-913">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bfb6a-914">パラメーター</span><span class="sxs-lookup"><span data-stu-id="bfb6a-914">Parameters</span></span>

|<span data-ttu-id="bfb6a-915">名前</span><span class="sxs-lookup"><span data-stu-id="bfb6a-915">Name</span></span>| <span data-ttu-id="bfb6a-916">型</span><span class="sxs-lookup"><span data-stu-id="bfb6a-916">Type</span></span>| <span data-ttu-id="bfb6a-917">属性</span><span class="sxs-lookup"><span data-stu-id="bfb6a-917">Attributes</span></span>| <span data-ttu-id="bfb6a-918">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-918">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="bfb6a-919">String</span><span class="sxs-lookup"><span data-stu-id="bfb6a-919">String</span></span>||<span data-ttu-id="bfb6a-920">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-920">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="bfb6a-921">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-921">Object</span></span>| <span data-ttu-id="bfb6a-922">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-922">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-923">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-923">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="bfb6a-924">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="bfb6a-924">Object</span></span>| <span data-ttu-id="bfb6a-925">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-925">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-926">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-926">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="bfb6a-927">function</span><span class="sxs-lookup"><span data-stu-id="bfb6a-927">function</span></span>| <span data-ttu-id="bfb6a-928">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="bfb6a-928">&lt;optional&gt;</span></span>|<span data-ttu-id="bfb6a-929">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-929">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="bfb6a-930">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-930">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bfb6a-931">エラー</span><span class="sxs-lookup"><span data-stu-id="bfb6a-931">Errors</span></span>

| <span data-ttu-id="bfb6a-932">エラー コード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-932">Error code</span></span> | <span data-ttu-id="bfb6a-933">説明</span><span class="sxs-lookup"><span data-stu-id="bfb6a-933">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="bfb6a-934">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-934">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bfb6a-935">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-935">Requirements</span></span>

|<span data-ttu-id="bfb6a-936">要件</span><span class="sxs-lookup"><span data-stu-id="bfb6a-936">Requirement</span></span>| <span data-ttu-id="bfb6a-937">値</span><span class="sxs-lookup"><span data-stu-id="bfb6a-937">Value</span></span>|
|---|---|
|[<span data-ttu-id="bfb6a-938">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bfb6a-938">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bfb6a-939">1.1</span><span class="sxs-lookup"><span data-stu-id="bfb6a-939">1.1</span></span>|
|[<span data-ttu-id="bfb6a-940">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bfb6a-940">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bfb6a-941">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="bfb6a-941">ReadWriteItem</span></span>|
|[<span data-ttu-id="bfb6a-942">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bfb6a-942">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bfb6a-943">作成</span><span class="sxs-lookup"><span data-stu-id="bfb6a-943">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="bfb6a-944">例</span><span class="sxs-lookup"><span data-stu-id="bfb6a-944">Example</span></span>

<span data-ttu-id="bfb6a-945">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="bfb6a-945">The following code removes an attachment with an identifier of '0'.</span></span>

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
