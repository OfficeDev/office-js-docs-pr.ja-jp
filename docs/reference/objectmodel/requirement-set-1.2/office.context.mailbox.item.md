---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7475e62c26d24ed9d191ca89934dd5d183b477fa
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696415"
---
# <a name="item"></a><span data-ttu-id="92deb-102">item</span><span class="sxs-lookup"><span data-stu-id="92deb-102">item</span></span>

### <span data-ttu-id="92deb-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="92deb-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="92deb-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-107">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-107">Requirements</span></span>

|<span data-ttu-id="92deb-108">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-108">Requirement</span></span>| <span data-ttu-id="92deb-109">値</span><span class="sxs-lookup"><span data-stu-id="92deb-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-111">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-111">1.0</span></span>|
|[<span data-ttu-id="92deb-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="92deb-113">Restricted</span></span>|
|[<span data-ttu-id="92deb-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="92deb-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-116">Members and methods</span></span>

| <span data-ttu-id="92deb-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="92deb-117">Member</span></span> | <span data-ttu-id="92deb-118">種類</span><span class="sxs-lookup"><span data-stu-id="92deb-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="92deb-119">attachments</span><span class="sxs-lookup"><span data-stu-id="92deb-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="92deb-120">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-120">Member</span></span> |
| [<span data-ttu-id="92deb-121">bcc</span><span class="sxs-lookup"><span data-stu-id="92deb-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="92deb-122">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-122">Member</span></span> |
| [<span data-ttu-id="92deb-123">body</span><span class="sxs-lookup"><span data-stu-id="92deb-123">body</span></span>](#body-body) | <span data-ttu-id="92deb-124">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-124">Member</span></span> |
| [<span data-ttu-id="92deb-125">cc</span><span class="sxs-lookup"><span data-stu-id="92deb-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="92deb-126">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-126">Member</span></span> |
| [<span data-ttu-id="92deb-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="92deb-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="92deb-128">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-128">Member</span></span> |
| [<span data-ttu-id="92deb-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="92deb-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="92deb-130">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-130">Member</span></span> |
| [<span data-ttu-id="92deb-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="92deb-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="92deb-132">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-132">Member</span></span> |
| [<span data-ttu-id="92deb-133">end</span><span class="sxs-lookup"><span data-stu-id="92deb-133">end</span></span>](#end-datetime) | <span data-ttu-id="92deb-134">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-134">Member</span></span> |
| [<span data-ttu-id="92deb-135">from</span><span class="sxs-lookup"><span data-stu-id="92deb-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="92deb-136">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-136">Member</span></span> |
| [<span data-ttu-id="92deb-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="92deb-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="92deb-138">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-138">Member</span></span> |
| [<span data-ttu-id="92deb-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="92deb-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="92deb-140">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-140">Member</span></span> |
| [<span data-ttu-id="92deb-141">itemId</span><span class="sxs-lookup"><span data-stu-id="92deb-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="92deb-142">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-142">Member</span></span> |
| [<span data-ttu-id="92deb-143">itemType</span><span class="sxs-lookup"><span data-stu-id="92deb-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="92deb-144">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-144">Member</span></span> |
| [<span data-ttu-id="92deb-145">location</span><span class="sxs-lookup"><span data-stu-id="92deb-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="92deb-146">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-146">Member</span></span> |
| [<span data-ttu-id="92deb-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="92deb-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="92deb-148">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-148">Member</span></span> |
| [<span data-ttu-id="92deb-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="92deb-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="92deb-150">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-150">Member</span></span> |
| [<span data-ttu-id="92deb-151">organizer</span><span class="sxs-lookup"><span data-stu-id="92deb-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="92deb-152">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-152">Member</span></span> |
| [<span data-ttu-id="92deb-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="92deb-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="92deb-154">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-154">Member</span></span> |
| [<span data-ttu-id="92deb-155">sender</span><span class="sxs-lookup"><span data-stu-id="92deb-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="92deb-156">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-156">Member</span></span> |
| [<span data-ttu-id="92deb-157">start</span><span class="sxs-lookup"><span data-stu-id="92deb-157">start</span></span>](#start-datetime) | <span data-ttu-id="92deb-158">Member</span><span class="sxs-lookup"><span data-stu-id="92deb-158">Member</span></span> |
| [<span data-ttu-id="92deb-159">subject</span><span class="sxs-lookup"><span data-stu-id="92deb-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="92deb-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="92deb-160">Member</span></span> |
| [<span data-ttu-id="92deb-161">to</span><span class="sxs-lookup"><span data-stu-id="92deb-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="92deb-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="92deb-162">Member</span></span> |
| [<span data-ttu-id="92deb-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="92deb-164">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-164">Method</span></span> |
| [<span data-ttu-id="92deb-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="92deb-166">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-166">Method</span></span> |
| [<span data-ttu-id="92deb-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="92deb-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="92deb-168">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-168">Method</span></span> |
| [<span data-ttu-id="92deb-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="92deb-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="92deb-170">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-170">Method</span></span> |
| [<span data-ttu-id="92deb-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="92deb-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="92deb-172">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-172">Method</span></span> |
| [<span data-ttu-id="92deb-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="92deb-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="92deb-174">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-174">Method</span></span> |
| [<span data-ttu-id="92deb-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="92deb-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="92deb-176">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-176">Method</span></span> |
| [<span data-ttu-id="92deb-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="92deb-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="92deb-178">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-178">Method</span></span> |
| [<span data-ttu-id="92deb-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="92deb-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="92deb-180">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-180">Method</span></span> |
| [<span data-ttu-id="92deb-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="92deb-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-182">Method</span></span> |
| [<span data-ttu-id="92deb-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="92deb-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-184">Method</span></span> |
| [<span data-ttu-id="92deb-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="92deb-186">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-186">Method</span></span> |
| [<span data-ttu-id="92deb-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="92deb-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="92deb-188">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="92deb-189">例</span><span class="sxs-lookup"><span data-stu-id="92deb-189">Example</span></span>

<span data-ttu-id="92deb-190">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="92deb-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="92deb-191">メンバー</span><span class="sxs-lookup"><span data-stu-id="92deb-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="92deb-192">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="92deb-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="92deb-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-195">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="92deb-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="92deb-196">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="92deb-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-197">型</span><span class="sxs-lookup"><span data-stu-id="92deb-197">Type</span></span>

*   <span data-ttu-id="92deb-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="92deb-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-199">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-199">Requirements</span></span>

|<span data-ttu-id="92deb-200">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-200">Requirement</span></span>| <span data-ttu-id="92deb-201">値</span><span class="sxs-lookup"><span data-stu-id="92deb-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-202">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-203">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-203">1.0</span></span>|
|[<span data-ttu-id="92deb-204">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-205">ReadItem</span></span>|
|[<span data-ttu-id="92deb-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-207">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-208">例</span><span class="sxs-lookup"><span data-stu-id="92deb-208">Example</span></span>

<span data-ttu-id="92deb-209">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="92deb-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="92deb-210">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-211">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="92deb-212">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-213">型</span><span class="sxs-lookup"><span data-stu-id="92deb-213">Type</span></span>

*   [<span data-ttu-id="92deb-214">受信者</span><span class="sxs-lookup"><span data-stu-id="92deb-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-215">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-215">Requirements</span></span>

|<span data-ttu-id="92deb-216">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-216">Requirement</span></span>| <span data-ttu-id="92deb-217">値</span><span class="sxs-lookup"><span data-stu-id="92deb-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-219">1.1</span><span class="sxs-lookup"><span data-stu-id="92deb-219">1.1</span></span>|
|[<span data-ttu-id="92deb-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-221">ReadItem</span></span>|
|[<span data-ttu-id="92deb-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-223">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-224">例</span><span class="sxs-lookup"><span data-stu-id="92deb-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="92deb-225">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-226">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-227">型</span><span class="sxs-lookup"><span data-stu-id="92deb-227">Type</span></span>

*   [<span data-ttu-id="92deb-228">Body</span><span class="sxs-lookup"><span data-stu-id="92deb-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-229">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-229">Requirements</span></span>

|<span data-ttu-id="92deb-230">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-230">Requirement</span></span>| <span data-ttu-id="92deb-231">値</span><span class="sxs-lookup"><span data-stu-id="92deb-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-232">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-233">1.1</span><span class="sxs-lookup"><span data-stu-id="92deb-233">1.1</span></span>|
|[<span data-ttu-id="92deb-234">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-235">ReadItem</span></span>|
|[<span data-ttu-id="92deb-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-238">例</span><span class="sxs-lookup"><span data-stu-id="92deb-238">Example</span></span>

<span data-ttu-id="92deb-239">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="92deb-240">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="92deb-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="92deb-241">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-242">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="92deb-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="92deb-243">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="92deb-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-244">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-244">Read mode</span></span>

<span data-ttu-id="92deb-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-247">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-247">Compose mode</span></span>

<span data-ttu-id="92deb-248">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92deb-249">型</span><span class="sxs-lookup"><span data-stu-id="92deb-249">Type</span></span>

*   <span data-ttu-id="92deb-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-251">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-251">Requirements</span></span>

|<span data-ttu-id="92deb-252">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-252">Requirement</span></span>| <span data-ttu-id="92deb-253">値</span><span class="sxs-lookup"><span data-stu-id="92deb-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-254">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-255">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-255">1.0</span></span>|
|[<span data-ttu-id="92deb-256">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-257">ReadItem</span></span>|
|[<span data-ttu-id="92deb-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-259">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-259">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="92deb-260">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="92deb-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="92deb-261">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="92deb-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="92deb-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="92deb-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-266">Type</span><span class="sxs-lookup"><span data-stu-id="92deb-266">Type</span></span>

*   <span data-ttu-id="92deb-267">String</span><span class="sxs-lookup"><span data-stu-id="92deb-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-268">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-268">Requirements</span></span>

|<span data-ttu-id="92deb-269">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-269">Requirement</span></span>| <span data-ttu-id="92deb-270">値</span><span class="sxs-lookup"><span data-stu-id="92deb-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-271">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-272">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-272">1.0</span></span>|
|[<span data-ttu-id="92deb-273">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-274">ReadItem</span></span>|
|[<span data-ttu-id="92deb-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-277">例</span><span class="sxs-lookup"><span data-stu-id="92deb-277">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="92deb-278">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="92deb-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="92deb-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-281">型</span><span class="sxs-lookup"><span data-stu-id="92deb-281">Type</span></span>

*   <span data-ttu-id="92deb-282">日付</span><span class="sxs-lookup"><span data-stu-id="92deb-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-283">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-283">Requirements</span></span>

|<span data-ttu-id="92deb-284">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-284">Requirement</span></span>| <span data-ttu-id="92deb-285">値</span><span class="sxs-lookup"><span data-stu-id="92deb-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-286">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-287">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-287">1.0</span></span>|
|[<span data-ttu-id="92deb-288">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-289">ReadItem</span></span>|
|[<span data-ttu-id="92deb-290">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-291">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-292">例</span><span class="sxs-lookup"><span data-stu-id="92deb-292">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="92deb-293">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="92deb-293">dateTimeModified: Date</span></span>

<span data-ttu-id="92deb-294">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="92deb-295">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="92deb-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-296">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-297">型</span><span class="sxs-lookup"><span data-stu-id="92deb-297">Type</span></span>

*   <span data-ttu-id="92deb-298">日付</span><span class="sxs-lookup"><span data-stu-id="92deb-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-299">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-299">Requirements</span></span>

|<span data-ttu-id="92deb-300">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-300">Requirement</span></span>| <span data-ttu-id="92deb-301">値</span><span class="sxs-lookup"><span data-stu-id="92deb-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-302">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-303">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-303">1.0</span></span>|
|[<span data-ttu-id="92deb-304">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-305">ReadItem</span></span>|
|[<span data-ttu-id="92deb-306">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-307">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-308">例</span><span class="sxs-lookup"><span data-stu-id="92deb-308">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="92deb-309">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-310">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="92deb-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-313">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-313">Read mode</span></span>

<span data-ttu-id="92deb-314">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-314">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-315">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-315">Compose mode</span></span>

<span data-ttu-id="92deb-316">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="92deb-317">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="92deb-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="92deb-318">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="92deb-319">型</span><span class="sxs-lookup"><span data-stu-id="92deb-319">Type</span></span>

*   <span data-ttu-id="92deb-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-321">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-321">Requirements</span></span>

|<span data-ttu-id="92deb-322">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-322">Requirement</span></span>| <span data-ttu-id="92deb-323">値</span><span class="sxs-lookup"><span data-stu-id="92deb-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-325">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-325">1.0</span></span>|
|[<span data-ttu-id="92deb-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-327">ReadItem</span></span>|
|[<span data-ttu-id="92deb-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-329">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-329">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="92deb-330">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="92deb-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-335">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="92deb-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-336">型</span><span class="sxs-lookup"><span data-stu-id="92deb-336">Type</span></span>

*   [<span data-ttu-id="92deb-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="92deb-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-338">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-338">Requirements</span></span>

|<span data-ttu-id="92deb-339">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-339">Requirement</span></span>| <span data-ttu-id="92deb-340">値</span><span class="sxs-lookup"><span data-stu-id="92deb-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-342">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-342">1.0</span></span>|
|[<span data-ttu-id="92deb-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-344">ReadItem</span></span>|
|[<span data-ttu-id="92deb-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-346">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-347">例</span><span class="sxs-lookup"><span data-stu-id="92deb-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="92deb-348">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="92deb-348">internetMessageId: String</span></span>

<span data-ttu-id="92deb-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-351">Type</span><span class="sxs-lookup"><span data-stu-id="92deb-351">Type</span></span>

*   <span data-ttu-id="92deb-352">String</span><span class="sxs-lookup"><span data-stu-id="92deb-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-353">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-353">Requirements</span></span>

|<span data-ttu-id="92deb-354">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-354">Requirement</span></span>| <span data-ttu-id="92deb-355">値</span><span class="sxs-lookup"><span data-stu-id="92deb-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-356">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-357">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-357">1.0</span></span>|
|[<span data-ttu-id="92deb-358">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-359">ReadItem</span></span>|
|[<span data-ttu-id="92deb-360">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-361">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-362">例</span><span class="sxs-lookup"><span data-stu-id="92deb-362">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="92deb-363">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="92deb-363">itemClass: String</span></span>

<span data-ttu-id="92deb-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="92deb-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="92deb-368">型</span><span class="sxs-lookup"><span data-stu-id="92deb-368">Type</span></span> | <span data-ttu-id="92deb-369">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-369">Description</span></span> | <span data-ttu-id="92deb-370">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="92deb-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="92deb-371">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="92deb-371">Appointment items</span></span> | <span data-ttu-id="92deb-372">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="92deb-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="92deb-373">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="92deb-373">Message items</span></span> | <span data-ttu-id="92deb-374">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="92deb-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="92deb-375">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-376">Type</span><span class="sxs-lookup"><span data-stu-id="92deb-376">Type</span></span>

*   <span data-ttu-id="92deb-377">String</span><span class="sxs-lookup"><span data-stu-id="92deb-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-378">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-378">Requirements</span></span>

|<span data-ttu-id="92deb-379">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-379">Requirement</span></span>| <span data-ttu-id="92deb-380">値</span><span class="sxs-lookup"><span data-stu-id="92deb-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-381">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-382">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-382">1.0</span></span>|
|[<span data-ttu-id="92deb-383">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-384">ReadItem</span></span>|
|[<span data-ttu-id="92deb-385">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-386">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-387">例</span><span class="sxs-lookup"><span data-stu-id="92deb-387">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="92deb-388">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="92deb-388">(nullable) itemId: String</span></span>

<span data-ttu-id="92deb-389">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="92deb-390">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-391">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="92deb-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="92deb-392">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="92deb-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="92deb-393">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="92deb-394">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="92deb-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-395">Type</span><span class="sxs-lookup"><span data-stu-id="92deb-395">Type</span></span>

*   <span data-ttu-id="92deb-396">String</span><span class="sxs-lookup"><span data-stu-id="92deb-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-397">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-397">Requirements</span></span>

|<span data-ttu-id="92deb-398">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-398">Requirement</span></span>| <span data-ttu-id="92deb-399">値</span><span class="sxs-lookup"><span data-stu-id="92deb-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-400">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-401">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-401">1.0</span></span>|
|[<span data-ttu-id="92deb-402">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-403">ReadItem</span></span>|
|[<span data-ttu-id="92deb-404">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-405">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-406">例</span><span class="sxs-lookup"><span data-stu-id="92deb-406">Example</span></span>

<span data-ttu-id="92deb-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="92deb-409">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-410">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="92deb-411">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="92deb-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-412">型</span><span class="sxs-lookup"><span data-stu-id="92deb-412">Type</span></span>

*   [<span data-ttu-id="92deb-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="92deb-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-414">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-414">Requirements</span></span>

|<span data-ttu-id="92deb-415">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-415">Requirement</span></span>| <span data-ttu-id="92deb-416">値</span><span class="sxs-lookup"><span data-stu-id="92deb-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-417">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-418">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-418">1.0</span></span>|
|[<span data-ttu-id="92deb-419">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-420">ReadItem</span></span>|
|[<span data-ttu-id="92deb-421">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-422">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-423">例</span><span class="sxs-lookup"><span data-stu-id="92deb-423">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="92deb-424">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-425">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-426">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-426">Read mode</span></span>

<span data-ttu-id="92deb-427">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-428">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-428">Compose mode</span></span>

<span data-ttu-id="92deb-429">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92deb-430">型</span><span class="sxs-lookup"><span data-stu-id="92deb-430">Type</span></span>

*   <span data-ttu-id="92deb-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-432">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-432">Requirements</span></span>

|<span data-ttu-id="92deb-433">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-433">Requirement</span></span>| <span data-ttu-id="92deb-434">値</span><span class="sxs-lookup"><span data-stu-id="92deb-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-435">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-436">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-436">1.0</span></span>|
|[<span data-ttu-id="92deb-437">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-438">ReadItem</span></span>|
|[<span data-ttu-id="92deb-439">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-440">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-440">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="92deb-441">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="92deb-441">normalizedSubject: String</span></span>

<span data-ttu-id="92deb-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="92deb-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-446">Type</span><span class="sxs-lookup"><span data-stu-id="92deb-446">Type</span></span>

*   <span data-ttu-id="92deb-447">String</span><span class="sxs-lookup"><span data-stu-id="92deb-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-448">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-448">Requirements</span></span>

|<span data-ttu-id="92deb-449">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-449">Requirement</span></span>| <span data-ttu-id="92deb-450">値</span><span class="sxs-lookup"><span data-stu-id="92deb-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-451">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-452">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-452">1.0</span></span>|
|[<span data-ttu-id="92deb-453">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-454">ReadItem</span></span>|
|[<span data-ttu-id="92deb-455">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-456">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-457">例</span><span class="sxs-lookup"><span data-stu-id="92deb-457">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="92deb-458">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-459">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="92deb-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="92deb-460">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="92deb-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-461">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-461">Read mode</span></span>

<span data-ttu-id="92deb-462">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-463">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-463">Compose mode</span></span>

<span data-ttu-id="92deb-464">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92deb-465">型</span><span class="sxs-lookup"><span data-stu-id="92deb-465">Type</span></span>

*   <span data-ttu-id="92deb-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-467">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-467">Requirements</span></span>

|<span data-ttu-id="92deb-468">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-468">Requirement</span></span>| <span data-ttu-id="92deb-469">値</span><span class="sxs-lookup"><span data-stu-id="92deb-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-470">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-471">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-471">1.0</span></span>|
|[<span data-ttu-id="92deb-472">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-473">ReadItem</span></span>|
|[<span data-ttu-id="92deb-474">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-475">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-475">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="92deb-476">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-479">型</span><span class="sxs-lookup"><span data-stu-id="92deb-479">Type</span></span>

*   [<span data-ttu-id="92deb-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="92deb-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-481">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-481">Requirements</span></span>

|<span data-ttu-id="92deb-482">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-482">Requirement</span></span>| <span data-ttu-id="92deb-483">値</span><span class="sxs-lookup"><span data-stu-id="92deb-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-485">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-485">1.0</span></span>|
|[<span data-ttu-id="92deb-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-487">ReadItem</span></span>|
|[<span data-ttu-id="92deb-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-489">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-490">例</span><span class="sxs-lookup"><span data-stu-id="92deb-490">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="92deb-491">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-492">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="92deb-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="92deb-493">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="92deb-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-494">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-494">Read mode</span></span>

<span data-ttu-id="92deb-495">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-496">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-496">Compose mode</span></span>

<span data-ttu-id="92deb-497">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="92deb-498">型</span><span class="sxs-lookup"><span data-stu-id="92deb-498">Type</span></span>

*   <span data-ttu-id="92deb-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-500">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-500">Requirements</span></span>

|<span data-ttu-id="92deb-501">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-501">Requirement</span></span>| <span data-ttu-id="92deb-502">値</span><span class="sxs-lookup"><span data-stu-id="92deb-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-504">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-504">1.0</span></span>|
|[<span data-ttu-id="92deb-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-506">ReadItem</span></span>|
|[<span data-ttu-id="92deb-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-508">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-508">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="92deb-509">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="92deb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="92deb-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-514">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="92deb-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="92deb-515">型</span><span class="sxs-lookup"><span data-stu-id="92deb-515">Type</span></span>

*   [<span data-ttu-id="92deb-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="92deb-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="92deb-517">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-517">Requirements</span></span>

|<span data-ttu-id="92deb-518">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-518">Requirement</span></span>| <span data-ttu-id="92deb-519">値</span><span class="sxs-lookup"><span data-stu-id="92deb-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-520">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-521">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-521">1.0</span></span>|
|[<span data-ttu-id="92deb-522">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-523">ReadItem</span></span>|
|[<span data-ttu-id="92deb-524">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-525">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-526">例</span><span class="sxs-lookup"><span data-stu-id="92deb-526">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="92deb-527">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-528">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="92deb-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-531">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-531">Read mode</span></span>

<span data-ttu-id="92deb-532">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-532">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-533">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-533">Compose mode</span></span>

<span data-ttu-id="92deb-534">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="92deb-535">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="92deb-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="92deb-536">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="92deb-537">型</span><span class="sxs-lookup"><span data-stu-id="92deb-537">Type</span></span>

*   <span data-ttu-id="92deb-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-539">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-539">Requirements</span></span>

|<span data-ttu-id="92deb-540">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-540">Requirement</span></span>| <span data-ttu-id="92deb-541">値</span><span class="sxs-lookup"><span data-stu-id="92deb-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-542">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-543">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-543">1.0</span></span>|
|[<span data-ttu-id="92deb-544">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-545">ReadItem</span></span>|
|[<span data-ttu-id="92deb-546">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-547">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-547">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="92deb-548">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-549">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="92deb-550">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="92deb-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-551">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-551">Read mode</span></span>

<span data-ttu-id="92deb-p130">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-554">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-554">Compose mode</span></span>

<span data-ttu-id="92deb-555">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="92deb-556">型</span><span class="sxs-lookup"><span data-stu-id="92deb-556">Type</span></span>

*   <span data-ttu-id="92deb-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-558">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-558">Requirements</span></span>

|<span data-ttu-id="92deb-559">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-559">Requirement</span></span>| <span data-ttu-id="92deb-560">値</span><span class="sxs-lookup"><span data-stu-id="92deb-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-561">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-562">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-562">1.0</span></span>|
|[<span data-ttu-id="92deb-563">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-564">ReadItem</span></span>|
|[<span data-ttu-id="92deb-565">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-566">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-566">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="92deb-567">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="92deb-568">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="92deb-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="92deb-569">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="92deb-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="92deb-570">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="92deb-570">Read mode</span></span>

<span data-ttu-id="92deb-p132">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="92deb-573">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="92deb-573">Compose mode</span></span>

<span data-ttu-id="92deb-574">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="92deb-575">型</span><span class="sxs-lookup"><span data-stu-id="92deb-575">Type</span></span>

*   <span data-ttu-id="92deb-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-577">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-577">Requirements</span></span>

|<span data-ttu-id="92deb-578">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-578">Requirement</span></span>| <span data-ttu-id="92deb-579">値</span><span class="sxs-lookup"><span data-stu-id="92deb-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-580">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-581">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-581">1.0</span></span>|
|[<span data-ttu-id="92deb-582">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-583">ReadItem</span></span>|
|[<span data-ttu-id="92deb-584">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-585">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="92deb-586">メソッド</span><span class="sxs-lookup"><span data-stu-id="92deb-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="92deb-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92deb-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="92deb-588">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="92deb-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="92deb-589">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="92deb-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="92deb-590">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-591">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-591">Parameters</span></span>

|<span data-ttu-id="92deb-592">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-592">Name</span></span>| <span data-ttu-id="92deb-593">種類</span><span class="sxs-lookup"><span data-stu-id="92deb-593">Type</span></span>| <span data-ttu-id="92deb-594">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-594">Attributes</span></span>| <span data-ttu-id="92deb-595">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="92deb-596">String</span><span class="sxs-lookup"><span data-stu-id="92deb-596">String</span></span>||<span data-ttu-id="92deb-p133">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="92deb-599">String</span><span class="sxs-lookup"><span data-stu-id="92deb-599">String</span></span>||<span data-ttu-id="92deb-p134">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="92deb-602">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-602">Object</span></span>| <span data-ttu-id="92deb-603">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-603">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-604">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="92deb-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="92deb-605">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-605">Object</span></span>| <span data-ttu-id="92deb-606">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-606">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-607">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="92deb-608">function</span><span class="sxs-lookup"><span data-stu-id="92deb-608">function</span></span>| <span data-ttu-id="92deb-609">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-609">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-610">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92deb-611">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="92deb-612">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="92deb-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92deb-613">エラー</span><span class="sxs-lookup"><span data-stu-id="92deb-613">Errors</span></span>

| <span data-ttu-id="92deb-614">エラー コード</span><span class="sxs-lookup"><span data-stu-id="92deb-614">Error code</span></span> | <span data-ttu-id="92deb-615">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="92deb-616">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="92deb-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="92deb-617">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="92deb-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="92deb-618">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="92deb-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-619">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-619">Requirements</span></span>

|<span data-ttu-id="92deb-620">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-620">Requirement</span></span>| <span data-ttu-id="92deb-621">値</span><span class="sxs-lookup"><span data-stu-id="92deb-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-622">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-623">1.1</span><span class="sxs-lookup"><span data-stu-id="92deb-623">1.1</span></span>|
|[<span data-ttu-id="92deb-624">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92deb-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="92deb-626">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-627">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-628">例</span><span class="sxs-lookup"><span data-stu-id="92deb-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="92deb-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92deb-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="92deb-630">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="92deb-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="92deb-p135">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="92deb-634">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="92deb-635">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-636">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-636">Parameters</span></span>

|<span data-ttu-id="92deb-637">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-637">Name</span></span>| <span data-ttu-id="92deb-638">種類</span><span class="sxs-lookup"><span data-stu-id="92deb-638">Type</span></span>| <span data-ttu-id="92deb-639">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-639">Attributes</span></span>| <span data-ttu-id="92deb-640">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="92deb-641">String</span><span class="sxs-lookup"><span data-stu-id="92deb-641">String</span></span>||<span data-ttu-id="92deb-p136">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="92deb-644">String</span><span class="sxs-lookup"><span data-stu-id="92deb-644">String</span></span>||<span data-ttu-id="92deb-645">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="92deb-645">The subject of the item to be attached.</span></span> <span data-ttu-id="92deb-646">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="92deb-647">Object</span><span class="sxs-lookup"><span data-stu-id="92deb-647">Object</span></span>| <span data-ttu-id="92deb-648">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-648">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-649">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="92deb-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="92deb-650">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-650">Object</span></span>| <span data-ttu-id="92deb-651">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-651">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-652">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="92deb-653">function</span><span class="sxs-lookup"><span data-stu-id="92deb-653">function</span></span>| <span data-ttu-id="92deb-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-654">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-655">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92deb-656">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="92deb-657">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="92deb-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92deb-658">エラー</span><span class="sxs-lookup"><span data-stu-id="92deb-658">Errors</span></span>

| <span data-ttu-id="92deb-659">エラー コード</span><span class="sxs-lookup"><span data-stu-id="92deb-659">Error code</span></span> | <span data-ttu-id="92deb-660">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="92deb-661">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="92deb-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-662">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-662">Requirements</span></span>

|<span data-ttu-id="92deb-663">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-663">Requirement</span></span>| <span data-ttu-id="92deb-664">値</span><span class="sxs-lookup"><span data-stu-id="92deb-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-665">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-666">1.1</span><span class="sxs-lookup"><span data-stu-id="92deb-666">1.1</span></span>|
|[<span data-ttu-id="92deb-667">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92deb-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="92deb-669">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-670">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-671">例</span><span class="sxs-lookup"><span data-stu-id="92deb-671">Example</span></span>

<span data-ttu-id="92deb-672">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="92deb-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="92deb-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="92deb-674">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-675">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="92deb-676">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="92deb-677">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="92deb-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="92deb-678">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="92deb-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="92deb-679">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="92deb-680">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="92deb-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-681">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-681">Parameters</span></span>

|<span data-ttu-id="92deb-682">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-682">Name</span></span>| <span data-ttu-id="92deb-683">型</span><span class="sxs-lookup"><span data-stu-id="92deb-683">Type</span></span>| <span data-ttu-id="92deb-684">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="92deb-685">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="92deb-685">String &#124; Object</span></span>| |<span data-ttu-id="92deb-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="92deb-688">**または**</span><span class="sxs-lookup"><span data-stu-id="92deb-688">**OR**</span></span><br/><span data-ttu-id="92deb-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="92deb-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="92deb-691">String</span><span class="sxs-lookup"><span data-stu-id="92deb-691">String</span></span> | <span data-ttu-id="92deb-692">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-692">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="92deb-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="92deb-696">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-696">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-697">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="92deb-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="92deb-698">String</span><span class="sxs-lookup"><span data-stu-id="92deb-698">String</span></span> | | <span data-ttu-id="92deb-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="92deb-701">String</span><span class="sxs-lookup"><span data-stu-id="92deb-701">String</span></span> | | <span data-ttu-id="92deb-702">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="92deb-703">文字列</span><span class="sxs-lookup"><span data-stu-id="92deb-703">String</span></span> | | <span data-ttu-id="92deb-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="92deb-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="92deb-706">String</span><span class="sxs-lookup"><span data-stu-id="92deb-706">String</span></span> | | <span data-ttu-id="92deb-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="92deb-710">function</span><span class="sxs-lookup"><span data-stu-id="92deb-710">function</span></span> | <span data-ttu-id="92deb-711">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-711">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-712">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-713">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-713">Requirements</span></span>

|<span data-ttu-id="92deb-714">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-714">Requirement</span></span>| <span data-ttu-id="92deb-715">値</span><span class="sxs-lookup"><span data-stu-id="92deb-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-716">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-717">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-717">1.0</span></span>|
|[<span data-ttu-id="92deb-718">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-719">ReadItem</span></span>|
|[<span data-ttu-id="92deb-720">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-721">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="92deb-722">例</span><span class="sxs-lookup"><span data-stu-id="92deb-722">Examples</span></span>

<span data-ttu-id="92deb-723">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="92deb-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="92deb-724">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-724">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="92deb-725">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-725">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="92deb-726">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="92deb-727">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="92deb-728">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="92deb-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="92deb-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="92deb-730">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-731">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="92deb-732">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="92deb-733">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="92deb-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="92deb-734">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="92deb-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="92deb-735">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="92deb-736">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="92deb-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-737">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-737">Parameters</span></span>

|<span data-ttu-id="92deb-738">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-738">Name</span></span>| <span data-ttu-id="92deb-739">型</span><span class="sxs-lookup"><span data-stu-id="92deb-739">Type</span></span>| <span data-ttu-id="92deb-740">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="92deb-741">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="92deb-741">String &#124; Object</span></span>| | <span data-ttu-id="92deb-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="92deb-744">**または**</span><span class="sxs-lookup"><span data-stu-id="92deb-744">**OR**</span></span><br/><span data-ttu-id="92deb-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="92deb-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="92deb-747">String</span><span class="sxs-lookup"><span data-stu-id="92deb-747">String</span></span> | <span data-ttu-id="92deb-748">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-748">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="92deb-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="92deb-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="92deb-752">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-752">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-753">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="92deb-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="92deb-754">String</span><span class="sxs-lookup"><span data-stu-id="92deb-754">String</span></span> | | <span data-ttu-id="92deb-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="92deb-757">String</span><span class="sxs-lookup"><span data-stu-id="92deb-757">String</span></span> | | <span data-ttu-id="92deb-758">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="92deb-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="92deb-759">文字列</span><span class="sxs-lookup"><span data-stu-id="92deb-759">String</span></span> | | <span data-ttu-id="92deb-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="92deb-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="92deb-762">String</span><span class="sxs-lookup"><span data-stu-id="92deb-762">String</span></span> | | <span data-ttu-id="92deb-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="92deb-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="92deb-766">function</span><span class="sxs-lookup"><span data-stu-id="92deb-766">function</span></span> | <span data-ttu-id="92deb-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-767">&lt;optional&gt;</span></span> | <span data-ttu-id="92deb-768">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-769">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-769">Requirements</span></span>

|<span data-ttu-id="92deb-770">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-770">Requirement</span></span>| <span data-ttu-id="92deb-771">値</span><span class="sxs-lookup"><span data-stu-id="92deb-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-772">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-773">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-773">1.0</span></span>|
|[<span data-ttu-id="92deb-774">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-775">ReadItem</span></span>|
|[<span data-ttu-id="92deb-776">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-777">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="92deb-778">例</span><span class="sxs-lookup"><span data-stu-id="92deb-778">Examples</span></span>

<span data-ttu-id="92deb-779">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="92deb-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="92deb-780">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-780">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="92deb-781">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-781">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="92deb-782">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="92deb-783">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="92deb-784">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="92deb-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="92deb-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="92deb-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="92deb-786">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-787">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-788">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-788">Requirements</span></span>

|<span data-ttu-id="92deb-789">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-789">Requirement</span></span>| <span data-ttu-id="92deb-790">値</span><span class="sxs-lookup"><span data-stu-id="92deb-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-791">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-792">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-792">1.0</span></span>|
|[<span data-ttu-id="92deb-793">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-794">ReadItem</span></span>|
|[<span data-ttu-id="92deb-795">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-796">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-797">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-797">Returns:</span></span>

<span data-ttu-id="92deb-798">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="92deb-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="92deb-799">例</span><span class="sxs-lookup"><span data-stu-id="92deb-799">Example</span></span>

<span data-ttu-id="92deb-800">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="92deb-800">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="92deb-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="92deb-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="92deb-802">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="92deb-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-803">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-804">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-804">Parameters</span></span>

|<span data-ttu-id="92deb-805">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-805">Name</span></span>| <span data-ttu-id="92deb-806">型</span><span class="sxs-lookup"><span data-stu-id="92deb-806">Type</span></span>| <span data-ttu-id="92deb-807">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="92deb-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="92deb-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="92deb-809">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="92deb-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92deb-810">Requirements</span><span class="sxs-lookup"><span data-stu-id="92deb-810">Requirements</span></span>

|<span data-ttu-id="92deb-811">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-811">Requirement</span></span>| <span data-ttu-id="92deb-812">値</span><span class="sxs-lookup"><span data-stu-id="92deb-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-813">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-814">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-814">1.0</span></span>|
|[<span data-ttu-id="92deb-815">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-816">制限あり</span><span class="sxs-lookup"><span data-stu-id="92deb-816">Restricted</span></span>|
|[<span data-ttu-id="92deb-817">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-818">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-819">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-819">Returns:</span></span>

<span data-ttu-id="92deb-820">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="92deb-821">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="92deb-822">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="92deb-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="92deb-823">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="92deb-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="92deb-824">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="92deb-824">Value of `entityType`</span></span> | <span data-ttu-id="92deb-825">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="92deb-825">Type of objects in returned array</span></span> | <span data-ttu-id="92deb-826">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="92deb-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="92deb-827">String</span><span class="sxs-lookup"><span data-stu-id="92deb-827">String</span></span> | <span data-ttu-id="92deb-828">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="92deb-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="92deb-829">連絡先</span><span class="sxs-lookup"><span data-stu-id="92deb-829">Contact</span></span> | <span data-ttu-id="92deb-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92deb-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="92deb-831">文字列</span><span class="sxs-lookup"><span data-stu-id="92deb-831">String</span></span> | <span data-ttu-id="92deb-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92deb-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="92deb-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="92deb-833">MeetingSuggestion</span></span> | <span data-ttu-id="92deb-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92deb-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="92deb-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="92deb-835">PhoneNumber</span></span> | <span data-ttu-id="92deb-836">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="92deb-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="92deb-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="92deb-837">TaskSuggestion</span></span> | <span data-ttu-id="92deb-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="92deb-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="92deb-839">文字列</span><span class="sxs-lookup"><span data-stu-id="92deb-839">String</span></span> | <span data-ttu-id="92deb-840">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="92deb-840">**Restricted**</span></span> |

<span data-ttu-id="92deb-841">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="92deb-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="92deb-842">例</span><span class="sxs-lookup"><span data-stu-id="92deb-842">Example</span></span>

<span data-ttu-id="92deb-843">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="92deb-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="92deb-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="92deb-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="92deb-845">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-846">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="92deb-847">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-848">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-848">Parameters</span></span>

|<span data-ttu-id="92deb-849">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-849">Name</span></span>| <span data-ttu-id="92deb-850">型</span><span class="sxs-lookup"><span data-stu-id="92deb-850">Type</span></span>| <span data-ttu-id="92deb-851">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="92deb-852">String</span><span class="sxs-lookup"><span data-stu-id="92deb-852">String</span></span>|<span data-ttu-id="92deb-853">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="92deb-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92deb-854">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-854">Requirements</span></span>

|<span data-ttu-id="92deb-855">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-855">Requirement</span></span>| <span data-ttu-id="92deb-856">値</span><span class="sxs-lookup"><span data-stu-id="92deb-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-857">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-858">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-858">1.0</span></span>|
|[<span data-ttu-id="92deb-859">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-860">ReadItem</span></span>|
|[<span data-ttu-id="92deb-861">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-862">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-863">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-863">Returns:</span></span>

<span data-ttu-id="92deb-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="92deb-866">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="92deb-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="92deb-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="92deb-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="92deb-868">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-869">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="92deb-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="92deb-873">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="92deb-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="92deb-874">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="92deb-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="92deb-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="92deb-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="92deb-877">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-877">Requirements</span></span>

|<span data-ttu-id="92deb-878">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-878">Requirement</span></span>| <span data-ttu-id="92deb-879">値</span><span class="sxs-lookup"><span data-stu-id="92deb-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-880">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-881">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-881">1.0</span></span>|
|[<span data-ttu-id="92deb-882">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-883">ReadItem</span></span>|
|[<span data-ttu-id="92deb-884">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-885">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-886">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-886">Returns:</span></span>

<span data-ttu-id="92deb-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="92deb-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="92deb-889">型: Object</span><span class="sxs-lookup"><span data-stu-id="92deb-889">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="92deb-890">例</span><span class="sxs-lookup"><span data-stu-id="92deb-890">Example</span></span>

<span data-ttu-id="92deb-891">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="92deb-891">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="92deb-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="92deb-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="92deb-893">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-893">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="92deb-894">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="92deb-894">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="92deb-895">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="92deb-895">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="92deb-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="92deb-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-898">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-898">Parameters</span></span>

|<span data-ttu-id="92deb-899">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-899">Name</span></span>| <span data-ttu-id="92deb-900">型</span><span class="sxs-lookup"><span data-stu-id="92deb-900">Type</span></span>| <span data-ttu-id="92deb-901">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-901">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="92deb-902">String</span><span class="sxs-lookup"><span data-stu-id="92deb-902">String</span></span>|<span data-ttu-id="92deb-903">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="92deb-903">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92deb-904">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-904">Requirements</span></span>

|<span data-ttu-id="92deb-905">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-905">Requirement</span></span>| <span data-ttu-id="92deb-906">値</span><span class="sxs-lookup"><span data-stu-id="92deb-906">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-907">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-907">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-908">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-908">1.0</span></span>|
|[<span data-ttu-id="92deb-909">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-909">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-910">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-910">ReadItem</span></span>|
|[<span data-ttu-id="92deb-911">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-911">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-912">読み取り</span><span class="sxs-lookup"><span data-stu-id="92deb-912">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-913">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-913">Returns:</span></span>

<span data-ttu-id="92deb-914">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="92deb-914">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="92deb-915">型: Array. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="92deb-915">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="92deb-916">例</span><span class="sxs-lookup"><span data-stu-id="92deb-916">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="92deb-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="92deb-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="92deb-918">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-918">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="92deb-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-921">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-921">Parameters</span></span>

|<span data-ttu-id="92deb-922">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-922">Name</span></span>| <span data-ttu-id="92deb-923">型</span><span class="sxs-lookup"><span data-stu-id="92deb-923">Type</span></span>| <span data-ttu-id="92deb-924">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-924">Attributes</span></span>| <span data-ttu-id="92deb-925">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-925">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="92deb-926">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="92deb-926">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="92deb-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="92deb-930">Object</span><span class="sxs-lookup"><span data-stu-id="92deb-930">Object</span></span>| <span data-ttu-id="92deb-931">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-931">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-932">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="92deb-932">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="92deb-933">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-933">Object</span></span>| <span data-ttu-id="92deb-934">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-934">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-935">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-935">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="92deb-936">function</span><span class="sxs-lookup"><span data-stu-id="92deb-936">function</span></span>||<span data-ttu-id="92deb-937">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92deb-938">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="92deb-938">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="92deb-939">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="92deb-939">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92deb-940">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-940">Requirements</span></span>

|<span data-ttu-id="92deb-941">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-941">Requirement</span></span>| <span data-ttu-id="92deb-942">値</span><span class="sxs-lookup"><span data-stu-id="92deb-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-943">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-944">1.2</span><span class="sxs-lookup"><span data-stu-id="92deb-944">1.2</span></span>|
|[<span data-ttu-id="92deb-945">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-946">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92deb-946">ReadWriteItem</span></span>|
|[<span data-ttu-id="92deb-947">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-948">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-948">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="92deb-949">戻り値:</span><span class="sxs-lookup"><span data-stu-id="92deb-949">Returns:</span></span>

<span data-ttu-id="92deb-950">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="92deb-950">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="92deb-951">型:String</span><span class="sxs-lookup"><span data-stu-id="92deb-951">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="92deb-952">例</span><span class="sxs-lookup"><span data-stu-id="92deb-952">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="92deb-953">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="92deb-953">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="92deb-954">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="92deb-954">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="92deb-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="92deb-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-958">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-958">Parameters</span></span>

|<span data-ttu-id="92deb-959">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-959">Name</span></span>| <span data-ttu-id="92deb-960">型</span><span class="sxs-lookup"><span data-stu-id="92deb-960">Type</span></span>| <span data-ttu-id="92deb-961">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-961">Attributes</span></span>| <span data-ttu-id="92deb-962">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-962">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="92deb-963">function</span><span class="sxs-lookup"><span data-stu-id="92deb-963">function</span></span>||<span data-ttu-id="92deb-964">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-964">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="92deb-965">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-965">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="92deb-966">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-966">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="92deb-967">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-967">Object</span></span>| <span data-ttu-id="92deb-968">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-968">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-969">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-969">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="92deb-970">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="92deb-970">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="92deb-971">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-971">Requirements</span></span>

|<span data-ttu-id="92deb-972">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-972">Requirement</span></span>| <span data-ttu-id="92deb-973">値</span><span class="sxs-lookup"><span data-stu-id="92deb-973">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-974">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-974">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-975">1.0</span><span class="sxs-lookup"><span data-stu-id="92deb-975">1.0</span></span>|
|[<span data-ttu-id="92deb-976">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-976">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-977">ReadItem</span><span class="sxs-lookup"><span data-stu-id="92deb-977">ReadItem</span></span>|
|[<span data-ttu-id="92deb-978">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-978">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-979">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="92deb-979">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-980">例</span><span class="sxs-lookup"><span data-stu-id="92deb-980">Example</span></span>

<span data-ttu-id="92deb-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="92deb-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="92deb-984">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="92deb-984">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="92deb-985">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="92deb-985">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="92deb-986">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="92deb-986">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="92deb-987">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="92deb-987">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="92deb-988">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="92deb-988">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="92deb-989">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="92deb-989">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-990">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-990">Parameters</span></span>

|<span data-ttu-id="92deb-991">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-991">Name</span></span>| <span data-ttu-id="92deb-992">型</span><span class="sxs-lookup"><span data-stu-id="92deb-992">Type</span></span>| <span data-ttu-id="92deb-993">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-993">Attributes</span></span>| <span data-ttu-id="92deb-994">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-994">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="92deb-995">String</span><span class="sxs-lookup"><span data-stu-id="92deb-995">String</span></span>||<span data-ttu-id="92deb-996">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="92deb-996">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="92deb-997">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-997">Object</span></span>| <span data-ttu-id="92deb-998">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-998">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-999">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="92deb-999">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="92deb-1000">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-1000">Object</span></span>| <span data-ttu-id="92deb-1001">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-1002">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1002">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="92deb-1003">関数</span><span class="sxs-lookup"><span data-stu-id="92deb-1003">function</span></span>| <span data-ttu-id="92deb-1004">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-1005">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1005">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="92deb-1006">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1006">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="92deb-1007">エラー</span><span class="sxs-lookup"><span data-stu-id="92deb-1007">Errors</span></span>

| <span data-ttu-id="92deb-1008">エラー コード</span><span class="sxs-lookup"><span data-stu-id="92deb-1008">Error code</span></span> | <span data-ttu-id="92deb-1009">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-1009">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="92deb-1010">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="92deb-1010">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-1011">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-1011">Requirements</span></span>

|<span data-ttu-id="92deb-1012">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-1012">Requirement</span></span>| <span data-ttu-id="92deb-1013">値</span><span class="sxs-lookup"><span data-stu-id="92deb-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-1014">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-1015">1.1</span><span class="sxs-lookup"><span data-stu-id="92deb-1015">1.1</span></span>|
|[<span data-ttu-id="92deb-1016">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92deb-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="92deb-1018">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-1019">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-1019">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-1020">例</span><span class="sxs-lookup"><span data-stu-id="92deb-1020">Example</span></span>

<span data-ttu-id="92deb-1021">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="92deb-1021">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="92deb-1022">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="92deb-1022">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="92deb-1023">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="92deb-1023">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="92deb-p166">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="92deb-1027">パラメーター</span><span class="sxs-lookup"><span data-stu-id="92deb-1027">Parameters</span></span>

|<span data-ttu-id="92deb-1028">名前</span><span class="sxs-lookup"><span data-stu-id="92deb-1028">Name</span></span>| <span data-ttu-id="92deb-1029">型</span><span class="sxs-lookup"><span data-stu-id="92deb-1029">Type</span></span>| <span data-ttu-id="92deb-1030">属性</span><span class="sxs-lookup"><span data-stu-id="92deb-1030">Attributes</span></span>| <span data-ttu-id="92deb-1031">説明</span><span class="sxs-lookup"><span data-stu-id="92deb-1031">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="92deb-1032">String</span><span class="sxs-lookup"><span data-stu-id="92deb-1032">String</span></span>||<span data-ttu-id="92deb-p167">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="92deb-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="92deb-1036">Object</span><span class="sxs-lookup"><span data-stu-id="92deb-1036">Object</span></span>| <span data-ttu-id="92deb-1037">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-1038">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="92deb-1038">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="92deb-1039">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="92deb-1039">Object</span></span>| <span data-ttu-id="92deb-1040">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-1041">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1041">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="92deb-1042">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="92deb-1042">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="92deb-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="92deb-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="92deb-1044">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1044">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="92deb-1045">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1045">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="92deb-1046">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1046">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="92deb-1047">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1047">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="92deb-1048">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="92deb-1048">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="92deb-1049">function</span><span class="sxs-lookup"><span data-stu-id="92deb-1049">function</span></span>||<span data-ttu-id="92deb-1050">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="92deb-1050">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="92deb-1051">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-1051">Requirements</span></span>

|<span data-ttu-id="92deb-1052">要件</span><span class="sxs-lookup"><span data-stu-id="92deb-1052">Requirement</span></span>| <span data-ttu-id="92deb-1053">値</span><span class="sxs-lookup"><span data-stu-id="92deb-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="92deb-1054">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="92deb-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="92deb-1055">1.2</span><span class="sxs-lookup"><span data-stu-id="92deb-1055">1.2</span></span>|
|[<span data-ttu-id="92deb-1056">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="92deb-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="92deb-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="92deb-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="92deb-1058">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="92deb-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="92deb-1059">作成</span><span class="sxs-lookup"><span data-stu-id="92deb-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="92deb-1060">例</span><span class="sxs-lookup"><span data-stu-id="92deb-1060">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
