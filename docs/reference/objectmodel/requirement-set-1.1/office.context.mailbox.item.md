---
title: Office. メールボックス-要件セット1.1
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 3d0b9783ea7fd235f4f989182ced04e0bce735ff
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682656"
---
# <a name="item"></a><span data-ttu-id="8ead0-102">item</span><span class="sxs-lookup"><span data-stu-id="8ead0-102">item</span></span>

### <span data-ttu-id="8ead0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="8ead0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="8ead0-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-107">Requirements</span></span>

|<span data-ttu-id="8ead0-108">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-108">Requirement</span></span>| <span data-ttu-id="8ead0-109">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-111">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-111">1.0</span></span>|
|[<span data-ttu-id="8ead0-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="8ead0-113">Restricted</span></span>|
|[<span data-ttu-id="8ead0-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8ead0-116">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="8ead0-116">Members and methods</span></span>

| <span data-ttu-id="8ead0-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="8ead0-117">Member</span></span> | <span data-ttu-id="8ead0-118">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8ead0-119">attachments</span><span class="sxs-lookup"><span data-stu-id="8ead0-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="8ead0-120">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-120">Member</span></span> |
| [<span data-ttu-id="8ead0-121">bcc</span><span class="sxs-lookup"><span data-stu-id="8ead0-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="8ead0-122">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-122">Member</span></span> |
| [<span data-ttu-id="8ead0-123">body</span><span class="sxs-lookup"><span data-stu-id="8ead0-123">body</span></span>](#body-body) | <span data-ttu-id="8ead0-124">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-124">Member</span></span> |
| [<span data-ttu-id="8ead0-125">cc</span><span class="sxs-lookup"><span data-stu-id="8ead0-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ead0-126">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-126">Member</span></span> |
| [<span data-ttu-id="8ead0-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="8ead0-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="8ead0-128">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-128">Member</span></span> |
| [<span data-ttu-id="8ead0-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="8ead0-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="8ead0-130">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-130">Member</span></span> |
| [<span data-ttu-id="8ead0-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="8ead0-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="8ead0-132">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-132">Member</span></span> |
| [<span data-ttu-id="8ead0-133">end</span><span class="sxs-lookup"><span data-stu-id="8ead0-133">end</span></span>](#end-datetime) | <span data-ttu-id="8ead0-134">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-134">Member</span></span> |
| [<span data-ttu-id="8ead0-135">from</span><span class="sxs-lookup"><span data-stu-id="8ead0-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="8ead0-136">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-136">Member</span></span> |
| [<span data-ttu-id="8ead0-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="8ead0-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="8ead0-138">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-138">Member</span></span> |
| [<span data-ttu-id="8ead0-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="8ead0-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="8ead0-140">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-140">Member</span></span> |
| [<span data-ttu-id="8ead0-141">itemId</span><span class="sxs-lookup"><span data-stu-id="8ead0-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="8ead0-142">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-142">Member</span></span> |
| [<span data-ttu-id="8ead0-143">itemType</span><span class="sxs-lookup"><span data-stu-id="8ead0-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="8ead0-144">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-144">Member</span></span> |
| [<span data-ttu-id="8ead0-145">location</span><span class="sxs-lookup"><span data-stu-id="8ead0-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="8ead0-146">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-146">Member</span></span> |
| [<span data-ttu-id="8ead0-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="8ead0-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="8ead0-148">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-148">Member</span></span> |
| [<span data-ttu-id="8ead0-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="8ead0-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ead0-150">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-150">Member</span></span> |
| [<span data-ttu-id="8ead0-151">organizer</span><span class="sxs-lookup"><span data-stu-id="8ead0-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="8ead0-152">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-152">Member</span></span> |
| [<span data-ttu-id="8ead0-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="8ead0-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ead0-154">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-154">Member</span></span> |
| [<span data-ttu-id="8ead0-155">sender</span><span class="sxs-lookup"><span data-stu-id="8ead0-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="8ead0-156">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-156">Member</span></span> |
| [<span data-ttu-id="8ead0-157">start</span><span class="sxs-lookup"><span data-stu-id="8ead0-157">start</span></span>](#start-datetime) | <span data-ttu-id="8ead0-158">Member</span><span class="sxs-lookup"><span data-stu-id="8ead0-158">Member</span></span> |
| [<span data-ttu-id="8ead0-159">subject</span><span class="sxs-lookup"><span data-stu-id="8ead0-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="8ead0-160">メンバー</span><span class="sxs-lookup"><span data-stu-id="8ead0-160">Member</span></span> |
| [<span data-ttu-id="8ead0-161">to</span><span class="sxs-lookup"><span data-stu-id="8ead0-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="8ead0-162">メンバー</span><span class="sxs-lookup"><span data-stu-id="8ead0-162">Member</span></span> |
| [<span data-ttu-id="8ead0-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ead0-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="8ead0-164">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-164">Method</span></span> |
| [<span data-ttu-id="8ead0-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ead0-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="8ead0-166">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-166">Method</span></span> |
| [<span data-ttu-id="8ead0-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="8ead0-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="8ead0-168">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-168">Method</span></span> |
| [<span data-ttu-id="8ead0-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="8ead0-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="8ead0-170">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-170">Method</span></span> |
| [<span data-ttu-id="8ead0-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="8ead0-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="8ead0-172">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-172">Method</span></span> |
| [<span data-ttu-id="8ead0-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="8ead0-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8ead0-174">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-174">Method</span></span> |
| [<span data-ttu-id="8ead0-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="8ead0-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="8ead0-176">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-176">Method</span></span> |
| [<span data-ttu-id="8ead0-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="8ead0-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="8ead0-178">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-178">Method</span></span> |
| [<span data-ttu-id="8ead0-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="8ead0-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="8ead0-180">Method</span><span class="sxs-lookup"><span data-stu-id="8ead0-180">Method</span></span> |
| [<span data-ttu-id="8ead0-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8ead0-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="8ead0-182">メソッド</span><span class="sxs-lookup"><span data-stu-id="8ead0-182">Method</span></span> |
| [<span data-ttu-id="8ead0-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="8ead0-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="8ead0-184">メソッド</span><span class="sxs-lookup"><span data-stu-id="8ead0-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="8ead0-185">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-185">Example</span></span>

<span data-ttu-id="8ead0-186">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8ead0-187">Members</span><span class="sxs-lookup"><span data-stu-id="8ead0-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="8ead0-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="8ead0-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="8ead0-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-191">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8ead0-192">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8ead0-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-193">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-193">Type</span></span>

*   <span data-ttu-id="8ead0-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="8ead0-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-195">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-195">Requirements</span></span>

|<span data-ttu-id="8ead0-196">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-196">Requirement</span></span>| <span data-ttu-id="8ead0-197">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-199">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-199">1.0</span></span>|
|[<span data-ttu-id="8ead0-200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-201">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-203">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-204">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-204">Example</span></span>

<span data-ttu-id="8ead0-205">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8ead0-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-207">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8ead0-208">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-208">Compose mode only.</span></span>

<span data-ttu-id="8ead0-209">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-209">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-210">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-210">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ead0-211">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-211">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ead0-212">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-212">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-213">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-213">Type</span></span>

*   [<span data-ttu-id="8ead0-214">受信者</span><span class="sxs-lookup"><span data-stu-id="8ead0-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-215">Requirements</span></span>

|<span data-ttu-id="8ead0-216">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-216">Requirement</span></span>| <span data-ttu-id="8ead0-217">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-219">1.1</span><span class="sxs-lookup"><span data-stu-id="8ead0-219">1.1</span></span>|
|[<span data-ttu-id="8ead0-220">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-221">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-222">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-223">作成</span><span class="sxs-lookup"><span data-stu-id="8ead0-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-224">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="8ead0-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-226">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-227">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-227">Type</span></span>

*   [<span data-ttu-id="8ead0-228">Body</span><span class="sxs-lookup"><span data-stu-id="8ead0-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-229">Requirements</span></span>

|<span data-ttu-id="8ead0-230">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-230">Requirement</span></span>| <span data-ttu-id="8ead0-231">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-232">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-233">1.1</span><span class="sxs-lookup"><span data-stu-id="8ead0-233">1.1</span></span>|
|[<span data-ttu-id="8ead0-234">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-235">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-238">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-238">Example</span></span>

<span data-ttu-id="8ead0-239">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8ead0-240">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8ead0-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-242">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8ead0-243">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-244">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-244">Read mode</span></span>

<span data-ttu-id="8ead0-245">`cc` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-245">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="8ead0-246">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-246">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-247">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-247">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-248">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-248">Compose mode</span></span>

<span data-ttu-id="8ead0-249">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="8ead0-250">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-251">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ead0-252">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ead0-253">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-254">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-254">Type</span></span>

*   <span data-ttu-id="8ead0-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-256">Requirements</span></span>

|<span data-ttu-id="8ead0-257">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-257">Requirement</span></span>| <span data-ttu-id="8ead0-258">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-259">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-260">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-260">1.0</span></span>|
|[<span data-ttu-id="8ead0-261">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-262">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8ead0-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="8ead0-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="8ead0-266">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8ead0-p110">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8ead0-p111">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-271">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-271">Type</span></span>

*   <span data-ttu-id="8ead0-272">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-273">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-273">Requirements</span></span>

|<span data-ttu-id="8ead0-274">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-274">Requirement</span></span>| <span data-ttu-id="8ead0-275">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-276">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-277">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-277">1.0</span></span>|
|[<span data-ttu-id="8ead0-278">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-279">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-280">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-281">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-282">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="8ead0-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="8ead0-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="8ead0-p112">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-286">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-286">Type</span></span>

*   <span data-ttu-id="8ead0-287">日付</span><span class="sxs-lookup"><span data-stu-id="8ead0-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-288">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-288">Requirements</span></span>

|<span data-ttu-id="8ead0-289">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-289">Requirement</span></span>| <span data-ttu-id="8ead0-290">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-291">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-292">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-292">1.0</span></span>|
|[<span data-ttu-id="8ead0-293">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-294">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-295">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-296">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-297">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="8ead0-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="8ead0-298">dateTimeModified: Date</span></span>

<span data-ttu-id="8ead0-p113">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-301">このメンバーは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-302">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-302">Type</span></span>

*   <span data-ttu-id="8ead0-303">日付</span><span class="sxs-lookup"><span data-stu-id="8ead0-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-304">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-304">Requirements</span></span>

|<span data-ttu-id="8ead0-305">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-305">Requirement</span></span>| <span data-ttu-id="8ead0-306">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-308">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-308">1.0</span></span>|
|[<span data-ttu-id="8ead0-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-310">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-313">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="8ead0-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-315">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8ead0-p114">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-318">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-318">Read mode</span></span>

<span data-ttu-id="8ead0-319">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-320">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-320">Compose mode</span></span>

<span data-ttu-id="8ead0-321">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8ead0-322">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8ead0-323">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8ead0-324">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-324">Type</span></span>

*   <span data-ttu-id="8ead0-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-326">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-326">Requirements</span></span>

|<span data-ttu-id="8ead0-327">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-327">Requirement</span></span>| <span data-ttu-id="8ead0-328">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-329">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-330">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-330">1.0</span></span>|
|[<span data-ttu-id="8ead0-331">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-332">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-333">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-334">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8ead0-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-p115">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8ead0-p116">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-340">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-341">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-341">Type</span></span>

*   [<span data-ttu-id="8ead0-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ead0-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-343">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-343">Requirements</span></span>

|<span data-ttu-id="8ead0-344">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-344">Requirement</span></span>| <span data-ttu-id="8ead0-345">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-346">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-347">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-347">1.0</span></span>|
|[<span data-ttu-id="8ead0-348">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-349">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-350">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-351">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-352">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="8ead0-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="8ead0-353">internetMessageId: String</span></span>

<span data-ttu-id="8ead0-p117">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-356">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-356">Type</span></span>

*   <span data-ttu-id="8ead0-357">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-358">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-358">Requirements</span></span>

|<span data-ttu-id="8ead0-359">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-359">Requirement</span></span>| <span data-ttu-id="8ead0-360">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-361">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-362">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-362">1.0</span></span>|
|[<span data-ttu-id="8ead0-363">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-364">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-365">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-366">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-367">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="8ead0-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="8ead0-368">itemClass: String</span></span>

<span data-ttu-id="8ead0-p118">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8ead0-p119">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8ead0-373">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-373">Type</span></span> | <span data-ttu-id="8ead0-374">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-374">Description</span></span> | <span data-ttu-id="8ead0-375">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="8ead0-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8ead0-376">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="8ead0-376">Appointment items</span></span> | <span data-ttu-id="8ead0-377">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8ead0-378">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="8ead0-378">Message items</span></span> | <span data-ttu-id="8ead0-379">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8ead0-380">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-381">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-381">Type</span></span>

*   <span data-ttu-id="8ead0-382">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-383">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-383">Requirements</span></span>

|<span data-ttu-id="8ead0-384">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-384">Requirement</span></span>| <span data-ttu-id="8ead0-385">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-386">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-387">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-387">1.0</span></span>|
|[<span data-ttu-id="8ead0-388">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-389">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-390">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-391">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-392">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8ead0-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="8ead0-393">(nullable) itemId: String</span></span>

<span data-ttu-id="8ead0-p120">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p120">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-396">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8ead0-397">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8ead0-398">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-398">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="8ead0-399">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8ead0-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-400">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-400">Type</span></span>

*   <span data-ttu-id="8ead0-401">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-402">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-402">Requirements</span></span>

|<span data-ttu-id="8ead0-403">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-403">Requirement</span></span>| <span data-ttu-id="8ead0-404">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-405">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-406">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-406">1.0</span></span>|
|[<span data-ttu-id="8ead0-407">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-407">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-408">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-409">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-409">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-410">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-411">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-411">Example</span></span>

<span data-ttu-id="8ead0-p122">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="8ead0-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-415">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-415">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8ead0-416">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-416">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-417">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-417">Type</span></span>

*   [<span data-ttu-id="8ead0-418">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8ead0-418">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-419">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-419">Requirements</span></span>

|<span data-ttu-id="8ead0-420">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-420">Requirement</span></span>| <span data-ttu-id="8ead0-421">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-421">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-422">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-423">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-423">1.0</span></span>|
|[<span data-ttu-id="8ead0-424">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-425">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-425">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-426">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-427">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-427">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-428">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-428">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="8ead0-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-430">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-430">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-431">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-431">Read mode</span></span>

<span data-ttu-id="8ead0-432">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-432">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-433">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-433">Compose mode</span></span>

<span data-ttu-id="8ead0-434">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-434">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-435">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-435">Type</span></span>

*   <span data-ttu-id="8ead0-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-437">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-437">Requirements</span></span>

|<span data-ttu-id="8ead0-438">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-438">Requirement</span></span>| <span data-ttu-id="8ead0-439">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-440">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-441">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-441">1.0</span></span>|
|[<span data-ttu-id="8ead0-442">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-443">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-444">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-445">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-445">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8ead0-446">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="8ead0-446">normalizedSubject: String</span></span>

<span data-ttu-id="8ead0-p123">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8ead0-p124">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-451">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-451">Type</span></span>

*   <span data-ttu-id="8ead0-452">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-453">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-453">Requirements</span></span>

|<span data-ttu-id="8ead0-454">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-454">Requirement</span></span>| <span data-ttu-id="8ead0-455">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-456">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-457">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-457">1.0</span></span>|
|[<span data-ttu-id="8ead0-458">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-459">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-460">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-461">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-462">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8ead0-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-464">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-464">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8ead0-465">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-465">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-466">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-466">Read mode</span></span>

<span data-ttu-id="8ead0-467">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-467">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="8ead0-468">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-468">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-469">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-469">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-470">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-470">Compose mode</span></span>

<span data-ttu-id="8ead0-471">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-471">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="8ead0-472">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-473">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-473">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ead0-474">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-474">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ead0-475">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-475">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-476">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-476">Type</span></span>

*   <span data-ttu-id="8ead0-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-478">Requirements</span></span>

|<span data-ttu-id="8ead0-479">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-479">Requirement</span></span>| <span data-ttu-id="8ead0-480">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-482">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-482">1.0</span></span>|
|[<span data-ttu-id="8ead0-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-484">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-486">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-486">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8ead0-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-p128">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-490">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-490">Type</span></span>

*   [<span data-ttu-id="8ead0-491">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ead0-491">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-492">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-492">Requirements</span></span>

|<span data-ttu-id="8ead0-493">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-493">Requirement</span></span>| <span data-ttu-id="8ead0-494">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-495">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-496">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-496">1.0</span></span>|
|[<span data-ttu-id="8ead0-497">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-498">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-499">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-500">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-500">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-501">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-501">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8ead0-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-503">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-503">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8ead0-504">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-504">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-505">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-505">Read mode</span></span>

<span data-ttu-id="8ead0-506">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-506">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="8ead0-507">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-507">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-508">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-508">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-509">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-509">Compose mode</span></span>

<span data-ttu-id="8ead0-510">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-510">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="8ead0-511">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-512">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-512">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ead0-513">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-513">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ead0-514">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-514">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-515">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-515">Type</span></span>

*   <span data-ttu-id="8ead0-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-517">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-517">Requirements</span></span>

|<span data-ttu-id="8ead0-518">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-518">Requirement</span></span>| <span data-ttu-id="8ead0-519">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-520">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-521">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-521">1.0</span></span>|
|[<span data-ttu-id="8ead0-522">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-523">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-524">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-525">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-525">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="8ead0-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-p132">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8ead0-p133">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-531">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-531">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8ead0-532">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-532">Type</span></span>

*   [<span data-ttu-id="8ead0-533">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8ead0-533">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="8ead0-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-534">Requirements</span></span>

|<span data-ttu-id="8ead0-535">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-535">Requirement</span></span>| <span data-ttu-id="8ead0-536">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-537">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-538">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-538">1.0</span></span>|
|[<span data-ttu-id="8ead0-539">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-540">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-541">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-542">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-542">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-543">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-543">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="8ead0-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-545">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-545">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8ead0-p134">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-548">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-548">Read mode</span></span>

<span data-ttu-id="8ead0-549">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-549">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-550">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-550">Compose mode</span></span>

<span data-ttu-id="8ead0-551">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-551">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8ead0-552">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-552">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8ead0-553">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-553">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8ead0-554">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-554">Type</span></span>

*   <span data-ttu-id="8ead0-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-556">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-556">Requirements</span></span>

|<span data-ttu-id="8ead0-557">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-557">Requirement</span></span>| <span data-ttu-id="8ead0-558">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-559">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-560">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-560">1.0</span></span>|
|[<span data-ttu-id="8ead0-561">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-562">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-563">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-564">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-564">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="8ead0-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-566">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-566">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8ead0-567">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-567">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-568">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-568">Read mode</span></span>

<span data-ttu-id="8ead0-p135">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-571">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-571">Compose mode</span></span>

<span data-ttu-id="8ead0-572">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-572">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-573">型</span><span class="sxs-lookup"><span data-stu-id="8ead0-573">Type</span></span>

*   <span data-ttu-id="8ead0-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-575">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-575">Requirements</span></span>

|<span data-ttu-id="8ead0-576">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-576">Requirement</span></span>| <span data-ttu-id="8ead0-577">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-577">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-578">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-578">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-579">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-579">1.0</span></span>|
|[<span data-ttu-id="8ead0-580">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-580">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-581">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-581">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-582">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-582">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-583">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-583">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="8ead0-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="8ead0-585">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-585">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8ead0-586">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-586">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8ead0-587">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-587">Read mode</span></span>

<span data-ttu-id="8ead0-588">`to` プロパティは、メッセージの `EmailAddressDetails` 行にある各受信者について、\*\*\*\* オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-588">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="8ead0-589">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-589">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-590">ただし、Windows および Mac では、最大500メンバーを取得することができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-590">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8ead0-591">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8ead0-591">Compose mode</span></span>

<span data-ttu-id="8ead0-592">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-592">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="8ead0-593">既定では、コレクションは最大100メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="8ead0-594">ただし、Windows と Mac では、次の制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-594">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="8ead0-595">500メンバーの最大数を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-595">Get 500 members maximum.</span></span>
- <span data-ttu-id="8ead0-596">1回の呼び出しで最大100のメンバーを設定します。最大数は500メンバーです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-596">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8ead0-597">Type</span><span class="sxs-lookup"><span data-stu-id="8ead0-597">Type</span></span>

*   <span data-ttu-id="8ead0-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-599">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-599">Requirements</span></span>

|<span data-ttu-id="8ead0-600">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-600">Requirement</span></span>| <span data-ttu-id="8ead0-601">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-602">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-603">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-603">1.0</span></span>|
|[<span data-ttu-id="8ead0-604">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-605">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-606">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-607">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-607">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8ead0-608">メソッド</span><span class="sxs-lookup"><span data-stu-id="8ead0-608">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8ead0-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ead0-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8ead0-610">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8ead0-611">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8ead0-612">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-613">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-613">Parameters</span></span>

|<span data-ttu-id="8ead0-614">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-614">Name</span></span>| <span data-ttu-id="8ead0-615">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-615">Type</span></span>| <span data-ttu-id="8ead0-616">属性</span><span class="sxs-lookup"><span data-stu-id="8ead0-616">Attributes</span></span>| <span data-ttu-id="8ead0-617">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8ead0-618">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-618">String</span></span>||<span data-ttu-id="8ead0-p139">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8ead0-621">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-621">String</span></span>||<span data-ttu-id="8ead0-p140">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8ead0-624">Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-624">Object</span></span>| <span data-ttu-id="8ead0-625">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-625">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-626">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8ead0-626">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ead0-627">Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-627">Object</span></span>| <span data-ttu-id="8ead0-628">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-628">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-629">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-629">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ead0-630">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-630">function</span></span>| <span data-ttu-id="8ead0-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-631">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-632">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ead0-633">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-633">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8ead0-634">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-634">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ead0-635">エラー</span><span class="sxs-lookup"><span data-stu-id="8ead0-635">Errors</span></span>

| <span data-ttu-id="8ead0-636">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8ead0-636">Error code</span></span> | <span data-ttu-id="8ead0-637">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-637">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8ead0-638">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-638">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8ead0-639">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-639">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8ead0-640">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-640">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ead0-641">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-641">Requirements</span></span>

|<span data-ttu-id="8ead0-642">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-642">Requirement</span></span>| <span data-ttu-id="8ead0-643">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-644">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-645">1.1</span><span class="sxs-lookup"><span data-stu-id="8ead0-645">1.1</span></span>|
|[<span data-ttu-id="8ead0-646">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-647">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-647">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ead0-648">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-649">作成</span><span class="sxs-lookup"><span data-stu-id="8ead0-649">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-650">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-650">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8ead0-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ead0-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8ead0-652">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-652">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8ead0-p141">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8ead0-656">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-656">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8ead0-657">Office アドインを Outlook on the web で実行している場合、編集中のアイテム以外のアイテムに `addItemAttachmentAsync` メソッドでアイテムを添付できます。ただし、これはサポートされていないため、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-657">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-658">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-658">Parameters</span></span>

|<span data-ttu-id="8ead0-659">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-659">Name</span></span>| <span data-ttu-id="8ead0-660">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-660">Type</span></span>| <span data-ttu-id="8ead0-661">属性</span><span class="sxs-lookup"><span data-stu-id="8ead0-661">Attributes</span></span>| <span data-ttu-id="8ead0-662">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-662">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8ead0-663">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-663">String</span></span>||<span data-ttu-id="8ead0-p142">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8ead0-666">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-666">String</span></span>||<span data-ttu-id="8ead0-667">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="8ead0-667">The subject of the item to be attached.</span></span> <span data-ttu-id="8ead0-668">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-668">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8ead0-669">Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-669">Object</span></span>| <span data-ttu-id="8ead0-670">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-670">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-671">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8ead0-671">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ead0-672">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8ead0-672">Object</span></span>| <span data-ttu-id="8ead0-673">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-673">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-674">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-674">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ead0-675">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-675">function</span></span>| <span data-ttu-id="8ead0-676">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-676">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-677">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ead0-678">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-678">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8ead0-679">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-679">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ead0-680">エラー</span><span class="sxs-lookup"><span data-stu-id="8ead0-680">Errors</span></span>

| <span data-ttu-id="8ead0-681">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8ead0-681">Error code</span></span> | <span data-ttu-id="8ead0-682">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-682">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8ead0-683">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-683">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ead0-684">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-684">Requirements</span></span>

|<span data-ttu-id="8ead0-685">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-685">Requirement</span></span>| <span data-ttu-id="8ead0-686">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-686">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-687">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-687">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-688">1.1</span><span class="sxs-lookup"><span data-stu-id="8ead0-688">1.1</span></span>|
|[<span data-ttu-id="8ead0-689">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-689">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-690">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-690">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ead0-691">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-691">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-692">作成</span><span class="sxs-lookup"><span data-stu-id="8ead0-692">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-693">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-693">Example</span></span>

<span data-ttu-id="8ead0-694">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-694">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8ead0-695">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8ead0-695">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8ead0-696">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-696">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-697">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-697">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ead0-698">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-698">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8ead0-699">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8ead0-699">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-700">へ`displayReplyAllForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-700">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8ead0-701">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="8ead0-701">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-702">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-702">Parameters</span></span>

|<span data-ttu-id="8ead0-703">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-703">Name</span></span>| <span data-ttu-id="8ead0-704">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-704">Type</span></span>| <span data-ttu-id="8ead0-705">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-705">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8ead0-706">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-706">String &#124; Object</span></span>| |<span data-ttu-id="8ead0-p145">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8ead0-709">**または**</span><span class="sxs-lookup"><span data-stu-id="8ead0-709">**OR**</span></span><br/><span data-ttu-id="8ead0-p146">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8ead0-712">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-712">String</span></span> | <span data-ttu-id="8ead0-713">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-713">&lt;optional&gt;</span></span> | <span data-ttu-id="8ead0-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8ead0-716">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-716">function</span></span> | <span data-ttu-id="8ead0-717">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-717">&lt;optional&gt;</span></span> | <span data-ttu-id="8ead0-718">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-718">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ead0-719">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-719">Requirements</span></span>

|<span data-ttu-id="8ead0-720">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-720">Requirement</span></span>| <span data-ttu-id="8ead0-721">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-721">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-722">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-722">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-723">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-723">1.0</span></span>|
|[<span data-ttu-id="8ead0-724">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-724">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-725">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-725">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-726">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-726">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-727">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-727">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8ead0-728">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-728">Examples</span></span>

<span data-ttu-id="8ead0-729">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-729">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8ead0-730">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-730">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8ead0-731">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-731">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8ead0-732">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-732">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8ead0-733">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8ead0-733">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8ead0-734">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-734">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-735">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-735">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ead0-736">Outlook on the web では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-736">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8ead0-737">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8ead0-737">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-738">へ`displayReplyForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-738">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8ead0-739">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="8ead0-739">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-740">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-740">Parameters</span></span>

|<span data-ttu-id="8ead0-741">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-741">Name</span></span>| <span data-ttu-id="8ead0-742">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-742">Type</span></span>| <span data-ttu-id="8ead0-743">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-743">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8ead0-744">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-744">String &#124; Object</span></span>| | <span data-ttu-id="8ead0-p149">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8ead0-747">**または**</span><span class="sxs-lookup"><span data-stu-id="8ead0-747">**OR**</span></span><br/><span data-ttu-id="8ead0-p150">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p150">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8ead0-750">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-750">String</span></span> | <span data-ttu-id="8ead0-751">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-751">&lt;optional&gt;</span></span> | <span data-ttu-id="8ead0-p151">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8ead0-754">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-754">function</span></span> | <span data-ttu-id="8ead0-755">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-755">&lt;optional&gt;</span></span> | <span data-ttu-id="8ead0-756">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ead0-757">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-757">Requirements</span></span>

|<span data-ttu-id="8ead0-758">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-758">Requirement</span></span>| <span data-ttu-id="8ead0-759">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-760">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-761">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-761">1.0</span></span>|
|[<span data-ttu-id="8ead0-762">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-763">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-764">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-765">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8ead0-766">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-766">Examples</span></span>

<span data-ttu-id="8ead0-767">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-767">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8ead0-768">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8ead0-769">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8ead0-770">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-770">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="8ead0-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="8ead0-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="8ead0-772">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-772">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-773">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-773">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-774">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-774">Requirements</span></span>

|<span data-ttu-id="8ead0-775">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-775">Requirement</span></span>| <span data-ttu-id="8ead0-776">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-776">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-777">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-777">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-778">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-778">1.0</span></span>|
|[<span data-ttu-id="8ead0-779">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-779">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-780">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-780">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-781">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-781">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-782">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-782">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ead0-783">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8ead0-783">Returns:</span></span>

<span data-ttu-id="8ead0-784">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="8ead0-784">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="8ead0-785">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-785">Example</span></span>

<span data-ttu-id="8ead0-786">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8ead0-786">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="8ead0-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="8ead0-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="8ead0-788">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-788">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-789">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-789">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-790">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-790">Parameters</span></span>

|<span data-ttu-id="8ead0-791">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-791">Name</span></span>| <span data-ttu-id="8ead0-792">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-792">Type</span></span>| <span data-ttu-id="8ead0-793">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-793">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8ead0-794">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8ead0-794">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="8ead0-795">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="8ead0-795">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ead0-796">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-796">Requirements</span></span>

|<span data-ttu-id="8ead0-797">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-797">Requirement</span></span>| <span data-ttu-id="8ead0-798">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-799">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-800">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-800">1.0</span></span>|
|[<span data-ttu-id="8ead0-801">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-802">制限あり</span><span class="sxs-lookup"><span data-stu-id="8ead0-802">Restricted</span></span>|
|[<span data-ttu-id="8ead0-803">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-804">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ead0-805">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8ead0-805">Returns:</span></span>

<span data-ttu-id="8ead0-806">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-806">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8ead0-807">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-807">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8ead0-808">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-808">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8ead0-809">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-809">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8ead0-810">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="8ead0-810">Value of `entityType`</span></span> | <span data-ttu-id="8ead0-811">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="8ead0-811">Type of objects in returned array</span></span> | <span data-ttu-id="8ead0-812">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-812">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8ead0-813">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-813">String</span></span> | <span data-ttu-id="8ead0-814">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8ead0-814">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8ead0-815">連絡先</span><span class="sxs-lookup"><span data-stu-id="8ead0-815">Contact</span></span> | <span data-ttu-id="8ead0-816">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ead0-816">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8ead0-817">文字列</span><span class="sxs-lookup"><span data-stu-id="8ead0-817">String</span></span> | <span data-ttu-id="8ead0-818">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ead0-818">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8ead0-819">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8ead0-819">MeetingSuggestion</span></span> | <span data-ttu-id="8ead0-820">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ead0-820">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8ead0-821">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8ead0-821">PhoneNumber</span></span> | <span data-ttu-id="8ead0-822">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8ead0-822">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8ead0-823">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8ead0-823">TaskSuggestion</span></span> | <span data-ttu-id="8ead0-824">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8ead0-824">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8ead0-825">文字列</span><span class="sxs-lookup"><span data-stu-id="8ead0-825">String</span></span> | <span data-ttu-id="8ead0-826">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8ead0-826">**Restricted**</span></span> |

<span data-ttu-id="8ead0-827">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="8ead0-827">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="8ead0-828">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-828">Example</span></span>

<span data-ttu-id="8ead0-829">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-829">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="8ead0-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="8ead0-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="8ead0-831">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-831">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-832">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-832">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ead0-833">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-833">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-834">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-834">Parameters</span></span>

|<span data-ttu-id="8ead0-835">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-835">Name</span></span>| <span data-ttu-id="8ead0-836">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-836">Type</span></span>| <span data-ttu-id="8ead0-837">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-837">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8ead0-838">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-838">String</span></span>|<span data-ttu-id="8ead0-839">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8ead0-839">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ead0-840">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-840">Requirements</span></span>

|<span data-ttu-id="8ead0-841">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-841">Requirement</span></span>| <span data-ttu-id="8ead0-842">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-843">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-844">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-844">1.0</span></span>|
|[<span data-ttu-id="8ead0-845">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-846">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-847">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-848">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ead0-849">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8ead0-849">Returns:</span></span>

<span data-ttu-id="8ead0-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="8ead0-852">型:Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="8ead0-852">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="8ead0-853">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8ead0-853">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8ead0-854">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-854">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-855">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ead0-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8ead0-859">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8ead0-859">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8ead0-860">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-860">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="8ead0-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8ead0-863">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-863">Requirements</span></span>

|<span data-ttu-id="8ead0-864">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-864">Requirement</span></span>| <span data-ttu-id="8ead0-865">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-867">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-867">1.0</span></span>|
|[<span data-ttu-id="8ead0-868">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-869">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-871">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ead0-872">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8ead0-872">Returns:</span></span>

<span data-ttu-id="8ead0-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="8ead0-875">型: Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-875">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="8ead0-876">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-876">Example</span></span>

<span data-ttu-id="8ead0-877">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="8ead0-877">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8ead0-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8ead0-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8ead0-879">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-879">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8ead0-880">このメソッドは、Outlook on iOS または Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-880">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8ead0-881">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-881">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8ead0-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-884">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-884">Parameters</span></span>

|<span data-ttu-id="8ead0-885">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-885">Name</span></span>| <span data-ttu-id="8ead0-886">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-886">Type</span></span>| <span data-ttu-id="8ead0-887">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8ead0-888">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-888">String</span></span>|<span data-ttu-id="8ead0-889">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8ead0-889">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ead0-890">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-890">Requirements</span></span>

|<span data-ttu-id="8ead0-891">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-891">Requirement</span></span>| <span data-ttu-id="8ead0-892">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-893">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-894">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-894">1.0</span></span>|
|[<span data-ttu-id="8ead0-895">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-896">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-897">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-898">読み取り</span><span class="sxs-lookup"><span data-stu-id="8ead0-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8ead0-899">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8ead0-899">Returns:</span></span>

<span data-ttu-id="8ead0-900">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="8ead0-900">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="8ead0-901">型: Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8ead0-901">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="8ead0-902">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-902">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8ead0-903">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8ead0-903">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8ead0-904">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-904">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8ead0-p158">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p158">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-908">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-908">Parameters</span></span>

|<span data-ttu-id="8ead0-909">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-909">Name</span></span>| <span data-ttu-id="8ead0-910">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-910">Type</span></span>| <span data-ttu-id="8ead0-911">属性</span><span class="sxs-lookup"><span data-stu-id="8ead0-911">Attributes</span></span>| <span data-ttu-id="8ead0-912">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-912">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8ead0-913">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-913">function</span></span>||<span data-ttu-id="8ead0-914">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-914">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8ead0-915">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-915">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8ead0-916">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-916">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8ead0-917">Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-917">Object</span></span>| <span data-ttu-id="8ead0-918">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-918">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-919">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-919">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8ead0-920">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-920">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8ead0-921">Requirements</span><span class="sxs-lookup"><span data-stu-id="8ead0-921">Requirements</span></span>

|<span data-ttu-id="8ead0-922">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-922">Requirement</span></span>| <span data-ttu-id="8ead0-923">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-923">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-924">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-924">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-925">1.0</span><span class="sxs-lookup"><span data-stu-id="8ead0-925">1.0</span></span>|
|[<span data-ttu-id="8ead0-926">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-926">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-927">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-927">ReadItem</span></span>|
|[<span data-ttu-id="8ead0-928">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-928">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-929">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8ead0-929">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-930">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-930">Example</span></span>

<span data-ttu-id="8ead0-p161">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-p161">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8ead0-934">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8ead0-934">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8ead0-935">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-935">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8ead0-936">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-936">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8ead0-937">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8ead0-937">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8ead0-938">Outlook on the web とモバイル デバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="8ead0-938">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8ead0-939">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-939">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8ead0-940">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8ead0-940">Parameters</span></span>

|<span data-ttu-id="8ead0-941">名前</span><span class="sxs-lookup"><span data-stu-id="8ead0-941">Name</span></span>| <span data-ttu-id="8ead0-942">種類</span><span class="sxs-lookup"><span data-stu-id="8ead0-942">Type</span></span>| <span data-ttu-id="8ead0-943">属性</span><span class="sxs-lookup"><span data-stu-id="8ead0-943">Attributes</span></span>| <span data-ttu-id="8ead0-944">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-944">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8ead0-945">String</span><span class="sxs-lookup"><span data-stu-id="8ead0-945">String</span></span>||<span data-ttu-id="8ead0-946">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="8ead0-946">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8ead0-947">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8ead0-947">Object</span></span>| <span data-ttu-id="8ead0-948">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-948">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-949">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8ead0-949">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8ead0-950">Object</span><span class="sxs-lookup"><span data-stu-id="8ead0-950">Object</span></span>| <span data-ttu-id="8ead0-951">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-951">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-952">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-952">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8ead0-953">function</span><span class="sxs-lookup"><span data-stu-id="8ead0-953">function</span></span>| <span data-ttu-id="8ead0-954">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8ead0-954">&lt;optional&gt;</span></span>|<span data-ttu-id="8ead0-955">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-955">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8ead0-956">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8ead0-956">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8ead0-957">エラー</span><span class="sxs-lookup"><span data-stu-id="8ead0-957">Errors</span></span>

| <span data-ttu-id="8ead0-958">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8ead0-958">Error code</span></span> | <span data-ttu-id="8ead0-959">説明</span><span class="sxs-lookup"><span data-stu-id="8ead0-959">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8ead0-960">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="8ead0-960">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8ead0-961">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-961">Requirements</span></span>

|<span data-ttu-id="8ead0-962">要件</span><span class="sxs-lookup"><span data-stu-id="8ead0-962">Requirement</span></span>| <span data-ttu-id="8ead0-963">値</span><span class="sxs-lookup"><span data-stu-id="8ead0-963">Value</span></span>|
|---|---|
|[<span data-ttu-id="8ead0-964">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8ead0-964">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8ead0-965">1.1</span><span class="sxs-lookup"><span data-stu-id="8ead0-965">1.1</span></span>|
|[<span data-ttu-id="8ead0-966">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8ead0-966">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8ead0-967">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8ead0-967">ReadWriteItem</span></span>|
|[<span data-ttu-id="8ead0-968">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8ead0-968">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8ead0-969">作成</span><span class="sxs-lookup"><span data-stu-id="8ead0-969">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8ead0-970">例</span><span class="sxs-lookup"><span data-stu-id="8ead0-970">Example</span></span>

<span data-ttu-id="8ead0-971">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8ead0-971">The following code removes an attachment with an identifier of '0'.</span></span>

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
