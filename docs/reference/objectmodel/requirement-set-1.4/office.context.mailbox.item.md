---
title: Office. メールボックス-要件セット1.4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 050f6daf3354b4ab7b617de37477ba639ef97129
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064719"
---
# <a name="item"></a><span data-ttu-id="1b0f1-102">item</span><span class="sxs-lookup"><span data-stu-id="1b0f1-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1b0f1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1b0f1-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1b0f1-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-106">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-106">Requirements</span></span>

|<span data-ttu-id="1b0f1-107">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-107">Requirement</span></span>| <span data-ttu-id="1b0f1-108">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-110">1.0</span></span>|
|[<span data-ttu-id="1b0f1-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="1b0f1-112">Restricted</span></span>|
|[<span data-ttu-id="1b0f1-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="1b0f1-115">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-115">Example</span></span>

<span data-ttu-id="1b0f1-116">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1b0f1-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="1b0f1-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-118">添付ファイル: <[Attachmentdetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="1b0f1-118">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="1b0f1-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-121">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1b0f1-122">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-123">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-123">Type</span></span>

*   <span data-ttu-id="1b0f1-124">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="1b0f1-124">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-125">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-125">Requirements</span></span>

|<span data-ttu-id="1b0f1-126">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-126">Requirement</span></span>| <span data-ttu-id="1b0f1-127">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-128">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-129">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-129">1.0</span></span>|
|[<span data-ttu-id="1b0f1-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-131">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-134">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-134">Example</span></span>

<span data-ttu-id="1b0f1-135">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-136">bcc:[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-136">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-137">メッセージの Bcc (ブラインドカーボンコピー) 行を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1b0f1-138">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-139">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-139">Type</span></span>

*   [<span data-ttu-id="1b0f1-140">受信者</span><span class="sxs-lookup"><span data-stu-id="1b0f1-140">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-141">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-141">Requirements</span></span>

|<span data-ttu-id="1b0f1-142">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-142">Requirement</span></span>| <span data-ttu-id="1b0f1-143">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-145">1.1</span><span class="sxs-lookup"><span data-stu-id="1b0f1-145">1.1</span></span>|
|[<span data-ttu-id="1b0f1-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-147">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-149">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-150">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="1b0f1-151">本文:[本文](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-151">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-152">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-153">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-153">Type</span></span>

*   [<span data-ttu-id="1b0f1-154">Body</span><span class="sxs-lookup"><span data-stu-id="1b0f1-154">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-155">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-155">Requirements</span></span>

|<span data-ttu-id="1b0f1-156">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-156">Requirement</span></span>| <span data-ttu-id="1b0f1-157">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1b0f1-159">1.1</span></span>|
|[<span data-ttu-id="1b0f1-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-161">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-164">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-164">Example</span></span>

<span data-ttu-id="1b0f1-165">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1b0f1-166">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-167">cc: <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-167">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-168">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1b0f1-169">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-170">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-170">Read mode</span></span>

<span data-ttu-id="1b0f1-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-173">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-173">Compose mode</span></span>

<span data-ttu-id="1b0f1-174">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-175">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-175">Type</span></span>

*   <span data-ttu-id="1b0f1-176">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-176">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-177">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-177">Requirements</span></span>

|<span data-ttu-id="1b0f1-178">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-178">Requirement</span></span>| <span data-ttu-id="1b0f1-179">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-181">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-181">1.0</span></span>|
|[<span data-ttu-id="1b0f1-182">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-183">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-185">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="1b0f1-186">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-186">(nullable) conversationId: String</span></span>

<span data-ttu-id="1b0f1-187">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1b0f1-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1b0f1-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-192">Type</span><span class="sxs-lookup"><span data-stu-id="1b0f1-192">Type</span></span>

*   <span data-ttu-id="1b0f1-193">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-194">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-194">Requirements</span></span>

|<span data-ttu-id="1b0f1-195">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-195">Requirement</span></span>| <span data-ttu-id="1b0f1-196">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-198">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-198">1.0</span></span>|
|[<span data-ttu-id="1b0f1-199">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-200">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-203">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="1b0f1-204">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="1b0f1-204">dateTimeCreated: Date</span></span>

<span data-ttu-id="1b0f1-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-207">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-207">Type</span></span>

*   <span data-ttu-id="1b0f1-208">日付</span><span class="sxs-lookup"><span data-stu-id="1b0f1-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-209">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-209">Requirements</span></span>

|<span data-ttu-id="1b0f1-210">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-210">Requirement</span></span>| <span data-ttu-id="1b0f1-211">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-213">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-213">1.0</span></span>|
|[<span data-ttu-id="1b0f1-214">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-215">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-217">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-218">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1b0f1-219">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="1b0f1-219">dateTimeModified: Date</span></span>

<span data-ttu-id="1b0f1-220">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-220">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="1b0f1-221">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-221">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-222">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-222">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-223">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-223">Type</span></span>

*   <span data-ttu-id="1b0f1-224">日付</span><span class="sxs-lookup"><span data-stu-id="1b0f1-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-225">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-225">Requirements</span></span>

|<span data-ttu-id="1b0f1-226">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-226">Requirement</span></span>| <span data-ttu-id="1b0f1-227">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-228">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-229">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-229">1.0</span></span>|
|[<span data-ttu-id="1b0f1-230">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-231">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-232">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-233">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-234">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="1b0f1-235">終了: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-235">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-236">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1b0f1-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-239">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-239">Read mode</span></span>

<span data-ttu-id="1b0f1-240">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-241">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-241">Compose mode</span></span>

<span data-ttu-id="1b0f1-242">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1b0f1-243">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-243">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1b0f1-244">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1b0f1-245">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-245">Type</span></span>

*   <span data-ttu-id="1b0f1-246">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-246">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-247">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-247">Requirements</span></span>

|<span data-ttu-id="1b0f1-248">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-248">Requirement</span></span>| <span data-ttu-id="1b0f1-249">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-251">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-251">1.0</span></span>|
|[<span data-ttu-id="1b0f1-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-253">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-256">from: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-256">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1b0f1-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-261">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-262">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-262">Type</span></span>

*   [<span data-ttu-id="1b0f1-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1b0f1-263">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-264">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-264">Requirements</span></span>

|<span data-ttu-id="1b0f1-265">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-265">Requirement</span></span>| <span data-ttu-id="1b0f1-266">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-268">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-268">1.0</span></span>|
|[<span data-ttu-id="1b0f1-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-270">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-272">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-273">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="1b0f1-274">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-274">internetMessageId: String</span></span>

<span data-ttu-id="1b0f1-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-277">Type</span><span class="sxs-lookup"><span data-stu-id="1b0f1-277">Type</span></span>

*   <span data-ttu-id="1b0f1-278">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-279">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-279">Requirements</span></span>

|<span data-ttu-id="1b0f1-280">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-280">Requirement</span></span>| <span data-ttu-id="1b0f1-281">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-283">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-283">1.0</span></span>|
|[<span data-ttu-id="1b0f1-284">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-285">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-287">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-288">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1b0f1-289">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-289">itemClass: String</span></span>

<span data-ttu-id="1b0f1-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1b0f1-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1b0f1-294">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-294">Type</span></span> | <span data-ttu-id="1b0f1-295">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-295">Description</span></span> | <span data-ttu-id="1b0f1-296">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="1b0f1-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1b0f1-297">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="1b0f1-297">Appointment items</span></span> | <span data-ttu-id="1b0f1-298">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1b0f1-299">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="1b0f1-299">Message items</span></span> | <span data-ttu-id="1b0f1-300">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1b0f1-301">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-302">Type</span><span class="sxs-lookup"><span data-stu-id="1b0f1-302">Type</span></span>

*   <span data-ttu-id="1b0f1-303">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-304">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-304">Requirements</span></span>

|<span data-ttu-id="1b0f1-305">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-305">Requirement</span></span>| <span data-ttu-id="1b0f1-306">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-308">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-308">1.0</span></span>|
|[<span data-ttu-id="1b0f1-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-310">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-313">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1b0f1-314">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-314">(nullable) itemId: String</span></span>

<span data-ttu-id="1b0f1-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-317">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1b0f1-318">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1b0f1-319">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1b0f1-320">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1b0f1-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-323">Type</span><span class="sxs-lookup"><span data-stu-id="1b0f1-323">Type</span></span>

*   <span data-ttu-id="1b0f1-324">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-325">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-325">Requirements</span></span>

|<span data-ttu-id="1b0f1-326">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-326">Requirement</span></span>| <span data-ttu-id="1b0f1-327">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-329">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-329">1.0</span></span>|
|[<span data-ttu-id="1b0f1-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-331">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-333">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-334">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-334">Example</span></span>

<span data-ttu-id="1b0f1-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="1b0f1-337">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-337">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-338">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1b0f1-339">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-340">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-340">Type</span></span>

*   [<span data-ttu-id="1b0f1-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1b0f1-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-342">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-342">Requirements</span></span>

|<span data-ttu-id="1b0f1-343">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-343">Requirement</span></span>| <span data-ttu-id="1b0f1-344">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-345">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-346">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-346">1.0</span></span>|
|[<span data-ttu-id="1b0f1-347">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-348">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-349">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-350">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-351">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="1b0f1-352">場所: String |[場所](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-352">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-353">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-354">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-354">Read mode</span></span>

<span data-ttu-id="1b0f1-355">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-356">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-356">Compose mode</span></span>

<span data-ttu-id="1b0f1-357">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-358">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-358">Type</span></span>

*   <span data-ttu-id="1b0f1-359">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-359">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-360">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-360">Requirements</span></span>

|<span data-ttu-id="1b0f1-361">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-361">Requirement</span></span>| <span data-ttu-id="1b0f1-362">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-364">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-364">1.0</span></span>|
|[<span data-ttu-id="1b0f1-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-366">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-368">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1b0f1-369">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-369">normalizedSubject: String</span></span>

<span data-ttu-id="1b0f1-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1b0f1-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-374">Type</span><span class="sxs-lookup"><span data-stu-id="1b0f1-374">Type</span></span>

*   <span data-ttu-id="1b0f1-375">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-376">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-376">Requirements</span></span>

|<span data-ttu-id="1b0f1-377">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-377">Requirement</span></span>| <span data-ttu-id="1b0f1-378">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-379">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-380">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-380">1.0</span></span>|
|[<span data-ttu-id="1b0f1-381">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-382">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-383">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-384">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-385">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="1b0f1-386">notificationMessages: [Notificationmessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-386">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-387">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-388">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-388">Type</span></span>

*   [<span data-ttu-id="1b0f1-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1b0f1-389">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-390">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-390">Requirements</span></span>

|<span data-ttu-id="1b0f1-391">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-391">Requirement</span></span>| <span data-ttu-id="1b0f1-392">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-393">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-394">1.3</span><span class="sxs-lookup"><span data-stu-id="1b0f1-394">1.3</span></span>|
|[<span data-ttu-id="1b0f1-395">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-396">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-397">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-398">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-399">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-400">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-400">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-401">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1b0f1-402">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-403">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-403">Read mode</span></span>

<span data-ttu-id="1b0f1-404">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-405">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-405">Compose mode</span></span>

<span data-ttu-id="1b0f1-406">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-407">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-407">Type</span></span>

*   <span data-ttu-id="1b0f1-408">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-408">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-409">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-409">Requirements</span></span>

|<span data-ttu-id="1b0f1-410">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-410">Requirement</span></span>| <span data-ttu-id="1b0f1-411">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-413">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-413">1.0</span></span>|
|[<span data-ttu-id="1b0f1-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-415">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-417">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-418">開催者: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-418">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-421">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-421">Type</span></span>

*   [<span data-ttu-id="1b0f1-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1b0f1-422">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-423">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-423">Requirements</span></span>

|<span data-ttu-id="1b0f1-424">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-424">Requirement</span></span>| <span data-ttu-id="1b0f1-425">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-427">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-427">1.0</span></span>|
|[<span data-ttu-id="1b0f1-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-429">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-431">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-432">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-433">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-433">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-434">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1b0f1-435">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-436">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-436">Read mode</span></span>

<span data-ttu-id="1b0f1-437">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-438">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-438">Compose mode</span></span>

<span data-ttu-id="1b0f1-439">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-440">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-440">Type</span></span>

*   <span data-ttu-id="1b0f1-441">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-441">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-442">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-442">Requirements</span></span>

|<span data-ttu-id="1b0f1-443">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-443">Requirement</span></span>| <span data-ttu-id="1b0f1-444">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-445">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-446">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-446">1.0</span></span>|
|[<span data-ttu-id="1b0f1-447">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-448">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-449">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-450">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-451">sender: [Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-451">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1b0f1-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-456">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1b0f1-457">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-457">Type</span></span>

*   [<span data-ttu-id="1b0f1-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1b0f1-458">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="1b0f1-459">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-459">Requirements</span></span>

|<span data-ttu-id="1b0f1-460">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-460">Requirement</span></span>| <span data-ttu-id="1b0f1-461">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-463">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-463">1.0</span></span>|
|[<span data-ttu-id="1b0f1-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-465">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-468">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="1b0f1-469">開始: 日付 |[時間](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-469">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-470">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1b0f1-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-473">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-473">Read mode</span></span>

<span data-ttu-id="1b0f1-474">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-475">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-475">Compose mode</span></span>

<span data-ttu-id="1b0f1-476">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1b0f1-477">[`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-477">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1b0f1-478">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1b0f1-479">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-479">Type</span></span>

*   <span data-ttu-id="1b0f1-480">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-480">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-481">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-481">Requirements</span></span>

|<span data-ttu-id="1b0f1-482">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-482">Requirement</span></span>| <span data-ttu-id="1b0f1-483">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-485">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-485">1.0</span></span>|
|[<span data-ttu-id="1b0f1-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-487">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-489">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-489">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="1b0f1-490">subject: String |[件名](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-490">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-491">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1b0f1-492">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-493">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-493">Read mode</span></span>

<span data-ttu-id="1b0f1-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-496">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-496">Compose mode</span></span>

<span data-ttu-id="1b0f1-497">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-498">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-498">Type</span></span>

*   <span data-ttu-id="1b0f1-499">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-499">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-500">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-500">Requirements</span></span>

|<span data-ttu-id="1b0f1-501">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-501">Requirement</span></span>| <span data-ttu-id="1b0f1-502">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-504">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-504">1.0</span></span>|
|[<span data-ttu-id="1b0f1-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-506">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-508">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-508">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="1b0f1-509">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)>|[受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-509">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="1b0f1-510">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1b0f1-511">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1b0f1-512">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-512">Read mode</span></span>

<span data-ttu-id="1b0f1-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1b0f1-515">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-515">Compose mode</span></span>

<span data-ttu-id="1b0f1-516">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1b0f1-517">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-517">Type</span></span>

*   <span data-ttu-id="1b0f1-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-519">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-519">Requirements</span></span>

|<span data-ttu-id="1b0f1-520">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-520">Requirement</span></span>| <span data-ttu-id="1b0f1-521">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-523">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-523">1.0</span></span>|
|[<span data-ttu-id="1b0f1-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-525">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-527">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1b0f1-528">メソッド</span><span class="sxs-lookup"><span data-stu-id="1b0f1-528">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1b0f1-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1b0f1-530">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1b0f1-531">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1b0f1-532">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-533">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-533">Parameters</span></span>

|<span data-ttu-id="1b0f1-534">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-534">Name</span></span>| <span data-ttu-id="1b0f1-535">種類</span><span class="sxs-lookup"><span data-stu-id="1b0f1-535">Type</span></span>| <span data-ttu-id="1b0f1-536">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-536">Attributes</span></span>| <span data-ttu-id="1b0f1-537">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1b0f1-538">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-538">String</span></span>||<span data-ttu-id="1b0f1-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1b0f1-541">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-541">String</span></span>||<span data-ttu-id="1b0f1-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1b0f1-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-544">Object</span></span>| <span data-ttu-id="1b0f1-545">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-545">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-546">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-547">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-547">Object</span></span>| <span data-ttu-id="1b0f1-548">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-548">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-549">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1b0f1-550">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-550">function</span></span>| <span data-ttu-id="1b0f1-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-551">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-552">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1b0f1-553">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1b0f1-554">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1b0f1-555">エラー</span><span class="sxs-lookup"><span data-stu-id="1b0f1-555">Errors</span></span>

| <span data-ttu-id="1b0f1-556">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-556">Error code</span></span> | <span data-ttu-id="1b0f1-557">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1b0f1-558">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1b0f1-559">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1b0f1-560">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-561">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-561">Requirements</span></span>

|<span data-ttu-id="1b0f1-562">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-562">Requirement</span></span>| <span data-ttu-id="1b0f1-563">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-565">1.1</span><span class="sxs-lookup"><span data-stu-id="1b0f1-565">1.1</span></span>|
|[<span data-ttu-id="1b0f1-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-569">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-570">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-570">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1b0f1-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1b0f1-572">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1b0f1-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1b0f1-576">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1b0f1-577">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-577">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-578">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-578">Parameters</span></span>

|<span data-ttu-id="1b0f1-579">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-579">Name</span></span>| <span data-ttu-id="1b0f1-580">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-580">Type</span></span>| <span data-ttu-id="1b0f1-581">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-581">Attributes</span></span>| <span data-ttu-id="1b0f1-582">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1b0f1-583">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-583">String</span></span>||<span data-ttu-id="1b0f1-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1b0f1-586">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-586">String</span></span>||<span data-ttu-id="1b0f1-587">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-587">The subject of the item to be attached.</span></span> <span data-ttu-id="1b0f1-588">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1b0f1-589">Object</span><span class="sxs-lookup"><span data-stu-id="1b0f1-589">Object</span></span>| <span data-ttu-id="1b0f1-590">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-590">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-591">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-592">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-592">Object</span></span>| <span data-ttu-id="1b0f1-593">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-593">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-594">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1b0f1-595">関数</span><span class="sxs-lookup"><span data-stu-id="1b0f1-595">function</span></span>| <span data-ttu-id="1b0f1-596">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-596">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1b0f1-598">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1b0f1-599">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1b0f1-600">エラー</span><span class="sxs-lookup"><span data-stu-id="1b0f1-600">Errors</span></span>

| <span data-ttu-id="1b0f1-601">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-601">Error code</span></span> | <span data-ttu-id="1b0f1-602">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1b0f1-603">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-604">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-604">Requirements</span></span>

|<span data-ttu-id="1b0f1-605">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-605">Requirement</span></span>| <span data-ttu-id="1b0f1-606">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-608">1.1</span><span class="sxs-lookup"><span data-stu-id="1b0f1-608">1.1</span></span>|
|[<span data-ttu-id="1b0f1-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-612">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-613">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-613">Example</span></span>

<span data-ttu-id="1b0f1-614">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="1b0f1-615">close()</span><span class="sxs-lookup"><span data-stu-id="1b0f1-615">close()</span></span>

<span data-ttu-id="1b0f1-616">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1b0f1-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-619">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1b0f1-620">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-621">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-621">Requirements</span></span>

|<span data-ttu-id="1b0f1-622">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-622">Requirement</span></span>| <span data-ttu-id="1b0f1-623">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-624">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-625">1.3</span><span class="sxs-lookup"><span data-stu-id="1b0f1-625">1.3</span></span>|
|[<span data-ttu-id="1b0f1-626">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-627">制限あり</span><span class="sxs-lookup"><span data-stu-id="1b0f1-627">Restricted</span></span>|
|[<span data-ttu-id="1b0f1-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-629">新規作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1b0f1-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1b0f1-631">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-632">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-632">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1b0f1-633">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-633">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1b0f1-634">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1b0f1-635">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-635">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="1b0f1-636">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-636">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="1b0f1-637">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-637">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-638">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-638">Parameters</span></span>

|<span data-ttu-id="1b0f1-639">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-639">Name</span></span>| <span data-ttu-id="1b0f1-640">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-640">Type</span></span>| <span data-ttu-id="1b0f1-641">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1b0f1-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1b0f1-642">String &#124; Object</span></span>| |<span data-ttu-id="1b0f1-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1b0f1-645">**または**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-645">**OR**</span></span><br/><span data-ttu-id="1b0f1-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1b0f1-648">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-648">String</span></span> | <span data-ttu-id="1b0f1-649">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-649">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1b0f1-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1b0f1-653">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-653">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-654">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1b0f1-655">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-655">String</span></span> | | <span data-ttu-id="1b0f1-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1b0f1-658">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-658">String</span></span> | | <span data-ttu-id="1b0f1-659">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1b0f1-660">文字列</span><span class="sxs-lookup"><span data-stu-id="1b0f1-660">String</span></span> | | <span data-ttu-id="1b0f1-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1b0f1-663">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-663">String</span></span> | | <span data-ttu-id="1b0f1-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1b0f1-667">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-667">function</span></span> | <span data-ttu-id="1b0f1-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-668">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-669">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-670">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-670">Requirements</span></span>

|<span data-ttu-id="1b0f1-671">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-671">Requirement</span></span>| <span data-ttu-id="1b0f1-672">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-673">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-674">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-674">1.0</span></span>|
|[<span data-ttu-id="1b0f1-675">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-676">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-678">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1b0f1-679">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-679">Examples</span></span>

<span data-ttu-id="1b0f1-680">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1b0f1-681">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1b0f1-682">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1b0f1-683">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1b0f1-684">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1b0f1-685">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1b0f1-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1b0f1-687">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-688">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-688">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1b0f1-689">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-689">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1b0f1-690">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1b0f1-691">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-691">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="1b0f1-692">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-692">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="1b0f1-693">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-693">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-694">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-694">Parameters</span></span>

|<span data-ttu-id="1b0f1-695">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-695">Name</span></span>| <span data-ttu-id="1b0f1-696">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-696">Type</span></span>| <span data-ttu-id="1b0f1-697">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1b0f1-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="1b0f1-698">String &#124; Object</span></span>| | <span data-ttu-id="1b0f1-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1b0f1-701">**または**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-701">**OR**</span></span><br/><span data-ttu-id="1b0f1-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1b0f1-704">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-704">String</span></span> | <span data-ttu-id="1b0f1-705">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-705">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1b0f1-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1b0f1-709">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-709">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-710">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1b0f1-711">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-711">String</span></span> | | <span data-ttu-id="1b0f1-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1b0f1-714">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-714">String</span></span> | | <span data-ttu-id="1b0f1-715">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1b0f1-716">文字列</span><span class="sxs-lookup"><span data-stu-id="1b0f1-716">String</span></span> | | <span data-ttu-id="1b0f1-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1b0f1-719">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-719">String</span></span> | | <span data-ttu-id="1b0f1-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1b0f1-723">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-723">function</span></span> | <span data-ttu-id="1b0f1-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-724">&lt;optional&gt;</span></span> | <span data-ttu-id="1b0f1-725">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-726">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-726">Requirements</span></span>

|<span data-ttu-id="1b0f1-727">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-727">Requirement</span></span>| <span data-ttu-id="1b0f1-728">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-729">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-730">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-730">1.0</span></span>|
|[<span data-ttu-id="1b0f1-731">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-732">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-734">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1b0f1-735">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-735">Examples</span></span>

<span data-ttu-id="1b0f1-736">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1b0f1-737">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1b0f1-738">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1b0f1-739">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1b0f1-740">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1b0f1-741">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="1b0f1-742">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-742">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="1b0f1-743">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-744">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-744">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-745">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-745">Requirements</span></span>

|<span data-ttu-id="1b0f1-746">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-746">Requirement</span></span>| <span data-ttu-id="1b0f1-747">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-749">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-749">1.0</span></span>|
|[<span data-ttu-id="1b0f1-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-751">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-753">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-754">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-754">Returns:</span></span>

<span data-ttu-id="1b0f1-755">型:[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-755">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="1b0f1-756">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-756">Example</span></span>

<span data-ttu-id="1b0f1-757">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="1b0f1-758">getEntitiesByType (entityType) > (nullable) {Array. < (String |[連絡先](/javascript/api/outlook/office.contact)|[会議の提案](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[tasksuggestion](/javascript/api/outlook/office.tasksuggestion)? view = outlook-js-1.4) >}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.4)>}</span></span>

<span data-ttu-id="1b0f1-759">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-760">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-760">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-761">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-761">Parameters</span></span>

|<span data-ttu-id="1b0f1-762">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-762">Name</span></span>| <span data-ttu-id="1b0f1-763">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-763">Type</span></span>| <span data-ttu-id="1b0f1-764">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1b0f1-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1b0f1-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="1b0f1-766">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-767">Requirements</span><span class="sxs-lookup"><span data-stu-id="1b0f1-767">Requirements</span></span>

|<span data-ttu-id="1b0f1-768">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-768">Requirement</span></span>| <span data-ttu-id="1b0f1-769">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-770">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-771">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-771">1.0</span></span>|
|[<span data-ttu-id="1b0f1-772">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-773">制限あり</span><span class="sxs-lookup"><span data-stu-id="1b0f1-773">Restricted</span></span>|
|[<span data-ttu-id="1b0f1-774">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-775">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-776">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-776">Returns:</span></span>

<span data-ttu-id="1b0f1-777">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1b0f1-778">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1b0f1-779">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1b0f1-780">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1b0f1-781">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-781">Value of `entityType`</span></span> | <span data-ttu-id="1b0f1-782">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-782">Type of objects in returned array</span></span> | <span data-ttu-id="1b0f1-783">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1b0f1-784">文字列</span><span class="sxs-lookup"><span data-stu-id="1b0f1-784">String</span></span> | <span data-ttu-id="1b0f1-785">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1b0f1-786">連絡先</span><span class="sxs-lookup"><span data-stu-id="1b0f1-786">Contact</span></span> | <span data-ttu-id="1b0f1-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1b0f1-788">文字列</span><span class="sxs-lookup"><span data-stu-id="1b0f1-788">String</span></span> | <span data-ttu-id="1b0f1-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1b0f1-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1b0f1-790">MeetingSuggestion</span></span> | <span data-ttu-id="1b0f1-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1b0f1-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1b0f1-792">PhoneNumber</span></span> | <span data-ttu-id="1b0f1-793">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1b0f1-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1b0f1-794">TaskSuggestion</span></span> | <span data-ttu-id="1b0f1-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1b0f1-796">文字列</span><span class="sxs-lookup"><span data-stu-id="1b0f1-796">String</span></span> | <span data-ttu-id="1b0f1-797">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="1b0f1-797">**Restricted**</span></span> |

<span data-ttu-id="1b0f1-798">型: < (文字列 |[連絡先](/javascript/api/outlook/office.contact)|[会議の提案](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[tasksuggestion](/javascript/api/outlook/office.tasksuggestion)? view = outlook-js-1.4) ></span><span class="sxs-lookup"><span data-stu-id="1b0f1-798">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.4)></span></span>

##### <a name="example"></a><span data-ttu-id="1b0f1-799">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-799">Example</span></span>

<span data-ttu-id="1b0f1-800">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="1b0f1-801">getFilteredEntitiesByName (name) > (nullable) {Array. < (String |[連絡先](/javascript/api/outlook/office.contact)|[会議の提案](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[tasksuggestion](/javascript/api/outlook/office.tasksuggestion)? view = outlook-js-1.4) >}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.4)>}</span></span>

<span data-ttu-id="1b0f1-802">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-803">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-803">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1b0f1-804">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-805">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-805">Parameters</span></span>

|<span data-ttu-id="1b0f1-806">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-806">Name</span></span>| <span data-ttu-id="1b0f1-807">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-807">Type</span></span>| <span data-ttu-id="1b0f1-808">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1b0f1-809">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-809">String</span></span>|<span data-ttu-id="1b0f1-810">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-811">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-811">Requirements</span></span>

|<span data-ttu-id="1b0f1-812">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-812">Requirement</span></span>| <span data-ttu-id="1b0f1-813">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-814">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-815">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-815">1.0</span></span>|
|[<span data-ttu-id="1b0f1-816">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-817">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-819">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-820">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-820">Returns:</span></span>

<span data-ttu-id="1b0f1-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1b0f1-823">型: < (文字列 |[連絡先](/javascript/api/outlook/office.contact)|[会議の提案](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[tasksuggestion](/javascript/api/outlook/office.tasksuggestion)? view = outlook-js-1.4) ></span><span class="sxs-lookup"><span data-stu-id="1b0f1-823">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.4)></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1b0f1-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1b0f1-825">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-826">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-826">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1b0f1-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1b0f1-830">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1b0f1-831">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1b0f1-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b0f1-835">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-835">Requirements</span></span>

|<span data-ttu-id="1b0f1-836">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-836">Requirement</span></span>| <span data-ttu-id="1b0f1-837">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-838">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-839">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-839">1.0</span></span>|
|[<span data-ttu-id="1b0f1-840">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-841">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-842">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-843">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-844">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-844">Returns:</span></span>

<span data-ttu-id="1b0f1-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1b0f1-847">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1b0f1-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1b0f1-848">Object</span><span class="sxs-lookup"><span data-stu-id="1b0f1-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1b0f1-849">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-849">Example</span></span>

<span data-ttu-id="1b0f1-850">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="1b0f1-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1b0f1-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1b0f1-852">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-853">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-853">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1b0f1-854">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1b0f1-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-857">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-857">Parameters</span></span>

|<span data-ttu-id="1b0f1-858">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-858">Name</span></span>| <span data-ttu-id="1b0f1-859">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-859">Type</span></span>| <span data-ttu-id="1b0f1-860">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1b0f1-861">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-861">String</span></span>|<span data-ttu-id="1b0f1-862">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-863">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-863">Requirements</span></span>

|<span data-ttu-id="1b0f1-864">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-864">Requirement</span></span>| <span data-ttu-id="1b0f1-865">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-867">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-867">1.0</span></span>|
|[<span data-ttu-id="1b0f1-868">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-869">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-871">読み取り</span><span class="sxs-lookup"><span data-stu-id="1b0f1-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-872">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-872">Returns:</span></span>

<span data-ttu-id="1b0f1-873">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1b0f1-874">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1b0f1-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1b0f1-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1b0f1-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1b0f1-876">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1b0f1-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1b0f1-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1b0f1-878">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1b0f1-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-881">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-881">Parameters</span></span>

|<span data-ttu-id="1b0f1-882">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-882">Name</span></span>| <span data-ttu-id="1b0f1-883">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-883">Type</span></span>| <span data-ttu-id="1b0f1-884">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-884">Attributes</span></span>| <span data-ttu-id="1b0f1-885">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1b0f1-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1b0f1-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1b0f1-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1b0f1-890">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-890">Object</span></span>| <span data-ttu-id="1b0f1-891">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-891">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-892">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-893">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-893">Object</span></span>| <span data-ttu-id="1b0f1-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-894">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-895">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1b0f1-896">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-896">function</span></span>||<span data-ttu-id="1b0f1-897">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1b0f1-898">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1b0f1-899">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-900">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-900">Requirements</span></span>

|<span data-ttu-id="1b0f1-901">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-901">Requirement</span></span>| <span data-ttu-id="1b0f1-902">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-903">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-904">1.2</span><span class="sxs-lookup"><span data-stu-id="1b0f1-904">1.2</span></span>|
|[<span data-ttu-id="1b0f1-905">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-908">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1b0f1-909">戻り値:</span><span class="sxs-lookup"><span data-stu-id="1b0f1-909">Returns:</span></span>

<span data-ttu-id="1b0f1-910">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1b0f1-911">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1b0f1-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1b0f1-912">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1b0f1-913">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-913">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1b0f1-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1b0f1-915">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1b0f1-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-919">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-919">Parameters</span></span>

|<span data-ttu-id="1b0f1-920">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-920">Name</span></span>| <span data-ttu-id="1b0f1-921">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-921">Type</span></span>| <span data-ttu-id="1b0f1-922">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-922">Attributes</span></span>| <span data-ttu-id="1b0f1-923">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1b0f1-924">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-924">function</span></span>||<span data-ttu-id="1b0f1-925">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1b0f1-926">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1b0f1-927">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1b0f1-928">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-928">Object</span></span>| <span data-ttu-id="1b0f1-929">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-929">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-930">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1b0f1-931">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-932">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-932">Requirements</span></span>

|<span data-ttu-id="1b0f1-933">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-933">Requirement</span></span>| <span data-ttu-id="1b0f1-934">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-936">1.0</span><span class="sxs-lookup"><span data-stu-id="1b0f1-936">1.0</span></span>|
|[<span data-ttu-id="1b0f1-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-938">ReadItem</span></span>|
|[<span data-ttu-id="1b0f1-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-940">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1b0f1-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-941">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-941">Example</span></span>

<span data-ttu-id="1b0f1-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1b0f1-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1b0f1-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1b0f1-946">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1b0f1-947">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-947">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="1b0f1-948">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-948">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="1b0f1-949">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-949">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="1b0f1-950">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-950">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-951">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-951">Parameters</span></span>

|<span data-ttu-id="1b0f1-952">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-952">Name</span></span>| <span data-ttu-id="1b0f1-953">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-953">Type</span></span>| <span data-ttu-id="1b0f1-954">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-954">Attributes</span></span>| <span data-ttu-id="1b0f1-955">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1b0f1-956">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-956">String</span></span>||<span data-ttu-id="1b0f1-957">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1b0f1-958">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-958">Object</span></span>| <span data-ttu-id="1b0f1-959">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-959">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-960">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-961">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-961">Object</span></span>| <span data-ttu-id="1b0f1-962">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-962">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-963">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1b0f1-964">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-964">function</span></span>| <span data-ttu-id="1b0f1-965">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-965">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-966">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1b0f1-967">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1b0f1-968">エラー</span><span class="sxs-lookup"><span data-stu-id="1b0f1-968">Errors</span></span>

| <span data-ttu-id="1b0f1-969">エラー コード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-969">Error code</span></span> | <span data-ttu-id="1b0f1-970">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1b0f1-971">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-972">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-972">Requirements</span></span>

|<span data-ttu-id="1b0f1-973">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-973">Requirement</span></span>| <span data-ttu-id="1b0f1-974">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-975">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-976">1.1</span><span class="sxs-lookup"><span data-stu-id="1b0f1-976">1.1</span></span>|
|[<span data-ttu-id="1b0f1-977">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-979">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-980">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-981">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-981">Example</span></span>

<span data-ttu-id="1b0f1-982">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="1b0f1-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="1b0f1-984">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="1b0f1-985">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-985">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="1b0f1-986">Outlook on the web または online モードの Outlook では、アイテムはサーバーに保存されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-986">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="1b0f1-987">キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-987">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-988">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1b0f1-989">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1b0f1-p168">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1b0f1-993">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1b0f1-994">Outlook on Mac では、会議の保存はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-994">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="1b0f1-995">新規`saveAsync`作成モードで会議から呼び出された場合、メソッドは失敗します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-995">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="1b0f1-996">回避策については[、「OFFICE JS API を使用して Outlook For Mac で会議を下書きとして保存できません](https://support.microsoft.com/help/4505745)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-996">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="1b0f1-997">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-998">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-998">Parameters</span></span>

|<span data-ttu-id="1b0f1-999">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-999">Name</span></span>| <span data-ttu-id="1b0f1-1000">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1000">Type</span></span>| <span data-ttu-id="1b0f1-1001">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1001">Attributes</span></span>| <span data-ttu-id="1b0f1-1002">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="1b0f1-1003">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1003">Object</span></span>| <span data-ttu-id="1b0f1-1004">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-1005">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-1006">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1006">Object</span></span>| <span data-ttu-id="1b0f1-1007">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-1008">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="1b0f1-1009">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1009">function</span></span>||<span data-ttu-id="1b0f1-1010">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1b0f1-1011">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1b0f1-1012">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1012">Requirements</span></span>

|<span data-ttu-id="1b0f1-1013">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1013">Requirement</span></span>| <span data-ttu-id="1b0f1-1014">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-1015">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1016">1.3</span></span>|
|[<span data-ttu-id="1b0f1-1017">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1017">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-1019">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1019">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-1020">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1b0f1-1021">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1021">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="1b0f1-p170">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1b0f1-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1b0f1-1025">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1b0f1-p171">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1b0f1-1029">パラメーター</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1029">Parameters</span></span>

|<span data-ttu-id="1b0f1-1030">名前</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1030">Name</span></span>| <span data-ttu-id="1b0f1-1031">型</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1031">Type</span></span>| <span data-ttu-id="1b0f1-1032">属性</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1032">Attributes</span></span>| <span data-ttu-id="1b0f1-1033">説明</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1b0f1-1034">String</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1034">String</span></span>||<span data-ttu-id="1b0f1-p172">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1b0f1-1038">Object</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1038">Object</span></span>| <span data-ttu-id="1b0f1-1039">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-1040">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1b0f1-1041">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1041">Object</span></span>| <span data-ttu-id="1b0f1-1042">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-1043">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1b0f1-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1b0f1-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="1b0f1-1046">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1046">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="1b0f1-1047">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1047">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1b0f1-1048">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1048">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="1b0f1-1049">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1049">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1b0f1-1050">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1b0f1-1051">function</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1051">function</span></span>||<span data-ttu-id="1b0f1-1052">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1b0f1-1053">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1053">Requirements</span></span>

|<span data-ttu-id="1b0f1-1054">要件</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1054">Requirement</span></span>| <span data-ttu-id="1b0f1-1055">値</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b0f1-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b0f1-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1057">1.2</span></span>|
|[<span data-ttu-id="1b0f1-1058">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b0f1-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="1b0f1-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b0f1-1061">作成</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1b0f1-1062">例</span><span class="sxs-lookup"><span data-stu-id="1b0f1-1062">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
