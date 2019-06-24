---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f0cf0e00a1bbd42b66b0b5e032599c54deb3ac6c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127437"
---
# <a name="item"></a><span data-ttu-id="8a62c-102">item</span><span class="sxs-lookup"><span data-stu-id="8a62c-102">item</span></span>

### <span data-ttu-id="8a62c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="8a62c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="8a62c-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-107">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-107">Requirements</span></span>

|<span data-ttu-id="8a62c-108">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-108">Requirement</span></span>| <span data-ttu-id="8a62c-109">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-111">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-111">1.0</span></span>|
|[<span data-ttu-id="8a62c-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="8a62c-113">Restricted</span></span>|
|[<span data-ttu-id="8a62c-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="8a62c-116">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-116">Example</span></span>

<span data-ttu-id="8a62c-117">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8a62c-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="8a62c-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="8a62c-119">添付ファイル: <[Attachmentdetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8a62c-119">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="8a62c-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-122">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8a62c-123">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8a62c-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-124">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-124">Type</span></span>

*   <span data-ttu-id="8a62c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8a62c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-126">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-126">Requirements</span></span>

|<span data-ttu-id="8a62c-127">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-127">Requirement</span></span>| <span data-ttu-id="8a62c-128">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-129">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-130">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-130">1.0</span></span>|
|[<span data-ttu-id="8a62c-131">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-132">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-134">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-135">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-135">Example</span></span>

<span data-ttu-id="8a62c-136">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="8a62c-137">bcc:[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-137">bcc: [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="8a62c-138">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8a62c-139">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-140">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-140">Type</span></span>

*   [<span data-ttu-id="8a62c-141">受信者</span><span class="sxs-lookup"><span data-stu-id="8a62c-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8a62c-142">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-142">Requirements</span></span>

|<span data-ttu-id="8a62c-143">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-143">Requirement</span></span>| <span data-ttu-id="8a62c-144">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-146">1.1</span><span class="sxs-lookup"><span data-stu-id="8a62c-146">1.1</span></span>|
|[<span data-ttu-id="8a62c-147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-148">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-150">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-151">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="8a62c-152">本文:[本文](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="8a62c-152">body: [Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="8a62c-153">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-154">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-154">Type</span></span>

*   [<span data-ttu-id="8a62c-155">Body</span><span class="sxs-lookup"><span data-stu-id="8a62c-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="8a62c-156">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-156">Requirements</span></span>

|<span data-ttu-id="8a62c-157">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-157">Requirement</span></span>| <span data-ttu-id="8a62c-158">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-160">1.1</span><span class="sxs-lookup"><span data-stu-id="8a62c-160">1.1</span></span>|
|[<span data-ttu-id="8a62c-161">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-162">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-165">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-165">Example</span></span>

<span data-ttu-id="8a62c-166">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="8a62c-167">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="8a62c-168">cc: <[emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-168">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="8a62c-169">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8a62c-170">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-171">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-171">Read mode</span></span>

<span data-ttu-id="8a62c-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-174">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-174">Compose mode</span></span>

<span data-ttu-id="8a62c-175">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-176">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-176">Type</span></span>

*   <span data-ttu-id="8a62c-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-178">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-178">Requirements</span></span>

|<span data-ttu-id="8a62c-179">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-179">Requirement</span></span>| <span data-ttu-id="8a62c-180">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-182">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-182">1.0</span></span>|
|[<span data-ttu-id="8a62c-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-184">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-186">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="8a62c-187">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="8a62c-187">(nullable) conversationId: String</span></span>

<span data-ttu-id="8a62c-188">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8a62c-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8a62c-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-193">Type</span><span class="sxs-lookup"><span data-stu-id="8a62c-193">Type</span></span>

*   <span data-ttu-id="8a62c-194">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-195">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-195">Requirements</span></span>

|<span data-ttu-id="8a62c-196">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-196">Requirement</span></span>| <span data-ttu-id="8a62c-197">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-199">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-199">1.0</span></span>|
|[<span data-ttu-id="8a62c-200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-201">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-204">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="8a62c-205">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="8a62c-205">dateTimeCreated: Date</span></span>

<span data-ttu-id="8a62c-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-208">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-208">Type</span></span>

*   <span data-ttu-id="8a62c-209">日付</span><span class="sxs-lookup"><span data-stu-id="8a62c-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-210">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-210">Requirements</span></span>

|<span data-ttu-id="8a62c-211">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-211">Requirement</span></span>| <span data-ttu-id="8a62c-212">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-214">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-214">1.0</span></span>|
|[<span data-ttu-id="8a62c-215">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-216">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-218">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-219">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8a62c-220">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="8a62c-220">dateTimeModified: Date</span></span>

<span data-ttu-id="8a62c-221">アイテムが最後に変更された日時を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-221">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="8a62c-222">Read mode only.</span><span class="sxs-lookup"><span data-stu-id="8a62c-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-223">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-223">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-224">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-224">Type</span></span>

*   <span data-ttu-id="8a62c-225">日付</span><span class="sxs-lookup"><span data-stu-id="8a62c-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-226">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-226">Requirements</span></span>

|<span data-ttu-id="8a62c-227">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-227">Requirement</span></span>| <span data-ttu-id="8a62c-228">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-229">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-230">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-230">1.0</span></span>|
|[<span data-ttu-id="8a62c-231">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-232">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-233">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-234">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-235">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="8a62c-236">終了: 日付 |[時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a62c-236">end: Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="8a62c-237">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8a62c-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-240">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-240">Read mode</span></span>

<span data-ttu-id="8a62c-241">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-242">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-242">Compose mode</span></span>

<span data-ttu-id="8a62c-243">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8a62c-244">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="8a62c-245">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8a62c-246">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-246">Type</span></span>

*   <span data-ttu-id="8a62c-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a62c-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-248">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-248">Requirements</span></span>

|<span data-ttu-id="8a62c-249">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-249">Requirement</span></span>| <span data-ttu-id="8a62c-250">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-252">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-252">1.0</span></span>|
|[<span data-ttu-id="8a62c-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-254">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-256">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="8a62c-257">from: [Emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8a62c-257">from: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="8a62c-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8a62c-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-262">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-263">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-263">Type</span></span>

*   [<span data-ttu-id="8a62c-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8a62c-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8a62c-265">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-265">Requirements</span></span>

|<span data-ttu-id="8a62c-266">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-266">Requirement</span></span>| <span data-ttu-id="8a62c-267">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-269">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-269">1.0</span></span>|
|[<span data-ttu-id="8a62c-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-271">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-273">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-274">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="8a62c-275">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="8a62c-275">internetMessageId: String</span></span>

<span data-ttu-id="8a62c-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-278">Type</span><span class="sxs-lookup"><span data-stu-id="8a62c-278">Type</span></span>

*   <span data-ttu-id="8a62c-279">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-280">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-280">Requirements</span></span>

|<span data-ttu-id="8a62c-281">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-281">Requirement</span></span>| <span data-ttu-id="8a62c-282">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-283">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-284">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-284">1.0</span></span>|
|[<span data-ttu-id="8a62c-285">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-286">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-288">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-289">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8a62c-290">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="8a62c-290">itemClass: String</span></span>

<span data-ttu-id="8a62c-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8a62c-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8a62c-295">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-295">Type</span></span> | <span data-ttu-id="8a62c-296">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-296">Description</span></span> | <span data-ttu-id="8a62c-297">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="8a62c-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8a62c-298">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="8a62c-298">Appointment items</span></span> | <span data-ttu-id="8a62c-299">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="8a62c-300">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="8a62c-300">Message items</span></span> | <span data-ttu-id="8a62c-301">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8a62c-302">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-303">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-303">Type</span></span>

*   <span data-ttu-id="8a62c-304">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-305">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-305">Requirements</span></span>

|<span data-ttu-id="8a62c-306">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-306">Requirement</span></span>| <span data-ttu-id="8a62c-307">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-309">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-309">1.0</span></span>|
|[<span data-ttu-id="8a62c-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-311">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-313">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-314">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8a62c-315">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="8a62c-315">(nullable) itemId: String</span></span>

<span data-ttu-id="8a62c-316">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-316">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="8a62c-317">閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-317">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-318">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8a62c-319">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8a62c-320">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="8a62c-321">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8a62c-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-322">Type</span><span class="sxs-lookup"><span data-stu-id="8a62c-322">Type</span></span>

*   <span data-ttu-id="8a62c-323">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-324">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-324">Requirements</span></span>

|<span data-ttu-id="8a62c-325">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-325">Requirement</span></span>| <span data-ttu-id="8a62c-326">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-328">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-328">1.0</span></span>|
|[<span data-ttu-id="8a62c-329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-330">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-332">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-333">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-333">Example</span></span>

<span data-ttu-id="8a62c-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="8a62c-336">itemType: [MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8a62c-336">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8a62c-337">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8a62c-338">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-339">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-339">Type</span></span>

*   [<span data-ttu-id="8a62c-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8a62c-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8a62c-341">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-341">Requirements</span></span>

|<span data-ttu-id="8a62c-342">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-342">Requirement</span></span>| <span data-ttu-id="8a62c-343">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-344">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-345">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-345">1.0</span></span>|
|[<span data-ttu-id="8a62c-346">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-347">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-348">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-349">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-350">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="8a62c-351">場所: String |[場所](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="8a62c-351">location: String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="8a62c-352">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-353">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-353">Read mode</span></span>

<span data-ttu-id="8a62c-354">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-355">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-355">Compose mode</span></span>

<span data-ttu-id="8a62c-356">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-357">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-357">Type</span></span>

*   <span data-ttu-id="8a62c-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="8a62c-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-359">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-359">Requirements</span></span>

|<span data-ttu-id="8a62c-360">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-360">Requirement</span></span>| <span data-ttu-id="8a62c-361">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-363">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-363">1.0</span></span>|
|[<span data-ttu-id="8a62c-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-365">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8a62c-368">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="8a62c-368">normalizedSubject: String</span></span>

<span data-ttu-id="8a62c-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8a62c-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-373">Type</span><span class="sxs-lookup"><span data-stu-id="8a62c-373">Type</span></span>

*   <span data-ttu-id="8a62c-374">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-375">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-375">Requirements</span></span>

|<span data-ttu-id="8a62c-376">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-376">Requirement</span></span>| <span data-ttu-id="8a62c-377">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-378">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-379">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-379">1.0</span></span>|
|[<span data-ttu-id="8a62c-380">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-381">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-382">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-383">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-384">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="8a62c-385">任意出席者: 配列. <[emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-385">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="8a62c-386">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8a62c-387">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-388">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-388">Read mode</span></span>

<span data-ttu-id="8a62c-389">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-390">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-390">Compose mode</span></span>

<span data-ttu-id="8a62c-391">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-392">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-392">Type</span></span>

*   <span data-ttu-id="8a62c-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-394">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-394">Requirements</span></span>

|<span data-ttu-id="8a62c-395">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-395">Requirement</span></span>| <span data-ttu-id="8a62c-396">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-397">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-398">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-398">1.0</span></span>|
|[<span data-ttu-id="8a62c-399">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-400">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-401">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-402">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="8a62c-403">開催者: [Emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8a62c-403">organizer: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="8a62c-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-406">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-406">Type</span></span>

*   [<span data-ttu-id="8a62c-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8a62c-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8a62c-408">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-408">Requirements</span></span>

|<span data-ttu-id="8a62c-409">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-409">Requirement</span></span>| <span data-ttu-id="8a62c-410">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-412">1.0</span></span>|
|[<span data-ttu-id="8a62c-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-414">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-417">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="8a62c-418">requiredatて dees: 配列. <[emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-418">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="8a62c-419">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8a62c-420">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-421">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-421">Read mode</span></span>

<span data-ttu-id="8a62c-422">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-423">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-423">Compose mode</span></span>

<span data-ttu-id="8a62c-424">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-425">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-425">Type</span></span>

*   <span data-ttu-id="8a62c-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-427">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-427">Requirements</span></span>

|<span data-ttu-id="8a62c-428">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-428">Requirement</span></span>| <span data-ttu-id="8a62c-429">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-431">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-431">1.0</span></span>|
|[<span data-ttu-id="8a62c-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-433">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-435">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="8a62c-436">sender: [Emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8a62c-436">sender: [EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="8a62c-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8a62c-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-441">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8a62c-442">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-442">Type</span></span>

*   [<span data-ttu-id="8a62c-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8a62c-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8a62c-444">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-444">Requirements</span></span>

|<span data-ttu-id="8a62c-445">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-445">Requirement</span></span>| <span data-ttu-id="8a62c-446">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-448">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-448">1.0</span></span>|
|[<span data-ttu-id="8a62c-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-450">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-453">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="8a62c-454">開始: 日付 |[時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a62c-454">start: Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="8a62c-455">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8a62c-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-458">Read mode</span></span>

<span data-ttu-id="8a62c-459">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-460">Compose mode</span></span>

<span data-ttu-id="8a62c-461">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8a62c-462">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="8a62c-463">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="8a62c-464">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-464">Type</span></span>

*   <span data-ttu-id="8a62c-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="8a62c-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-466">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-466">Requirements</span></span>

|<span data-ttu-id="8a62c-467">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-467">Requirement</span></span>| <span data-ttu-id="8a62c-468">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-470">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-470">1.0</span></span>|
|[<span data-ttu-id="8a62c-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-472">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-474">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-474">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="8a62c-475">subject: String |[件名](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8a62c-475">subject: String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="8a62c-476">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8a62c-477">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-478">Read mode</span></span>

<span data-ttu-id="8a62c-p130">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-481">Compose mode</span></span>

<span data-ttu-id="8a62c-482">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-483">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-483">Type</span></span>

*   <span data-ttu-id="8a62c-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8a62c-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-485">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-485">Requirements</span></span>

|<span data-ttu-id="8a62c-486">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-486">Requirement</span></span>| <span data-ttu-id="8a62c-487">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-489">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-489">1.0</span></span>|
|[<span data-ttu-id="8a62c-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-491">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-493">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-493">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="8a62c-494">宛先: 配列. <[emailaddressdetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-494">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="8a62c-495">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8a62c-496">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8a62c-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-497">Read mode</span></span>

<span data-ttu-id="8a62c-p132">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="8a62c-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="8a62c-500">Compose mode</span></span>

<span data-ttu-id="8a62c-501">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8a62c-502">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-502">Type</span></span>

*   <span data-ttu-id="8a62c-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8a62c-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-504">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-504">Requirements</span></span>

|<span data-ttu-id="8a62c-505">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-505">Requirement</span></span>| <span data-ttu-id="8a62c-506">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-508">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-508">1.0</span></span>|
|[<span data-ttu-id="8a62c-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-510">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-512">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8a62c-513">メソッド</span><span class="sxs-lookup"><span data-stu-id="8a62c-513">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8a62c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a62c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8a62c-515">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8a62c-516">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8a62c-517">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-518">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-518">Parameters</span></span>

|<span data-ttu-id="8a62c-519">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-519">Name</span></span>| <span data-ttu-id="8a62c-520">種類</span><span class="sxs-lookup"><span data-stu-id="8a62c-520">Type</span></span>| <span data-ttu-id="8a62c-521">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-521">Attributes</span></span>| <span data-ttu-id="8a62c-522">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8a62c-523">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-523">String</span></span>||<span data-ttu-id="8a62c-p133">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8a62c-526">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-526">String</span></span>||<span data-ttu-id="8a62c-p134">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8a62c-529">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-529">Object</span></span>| <span data-ttu-id="8a62c-530">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-530">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-531">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8a62c-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8a62c-532">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-532">Object</span></span>| <span data-ttu-id="8a62c-533">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-533">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-534">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8a62c-535">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-535">function</span></span>| <span data-ttu-id="8a62c-536">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-536">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-537">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a62c-538">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8a62c-539">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a62c-540">エラー</span><span class="sxs-lookup"><span data-stu-id="8a62c-540">Errors</span></span>

| <span data-ttu-id="8a62c-541">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8a62c-541">Error code</span></span> | <span data-ttu-id="8a62c-542">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8a62c-543">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8a62c-544">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8a62c-545">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-546">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-546">Requirements</span></span>

|<span data-ttu-id="8a62c-547">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-547">Requirement</span></span>| <span data-ttu-id="8a62c-548">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-550">1.1</span><span class="sxs-lookup"><span data-stu-id="8a62c-550">1.1</span></span>|
|[<span data-ttu-id="8a62c-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a62c-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-554">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-555">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-555">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8a62c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a62c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8a62c-557">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8a62c-p135">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8a62c-561">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8a62c-562">Office アドインが Outlook on the web で実行されている場合、 `addItemAttachmentAsync`メソッドは、編集しているアイテム以外のアイテムにアイテムを添付できます。ただし、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-562">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-563">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-563">Parameters</span></span>

|<span data-ttu-id="8a62c-564">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-564">Name</span></span>| <span data-ttu-id="8a62c-565">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-565">Type</span></span>| <span data-ttu-id="8a62c-566">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-566">Attributes</span></span>| <span data-ttu-id="8a62c-567">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8a62c-568">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-568">String</span></span>||<span data-ttu-id="8a62c-p136">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8a62c-571">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-571">String</span></span>||<span data-ttu-id="8a62c-572">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="8a62c-572">The subject of the item to be attached.</span></span> <span data-ttu-id="8a62c-573">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8a62c-574">Object</span><span class="sxs-lookup"><span data-stu-id="8a62c-574">Object</span></span>| <span data-ttu-id="8a62c-575">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-575">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-576">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8a62c-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8a62c-577">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-577">Object</span></span>| <span data-ttu-id="8a62c-578">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-578">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-579">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8a62c-580">関数</span><span class="sxs-lookup"><span data-stu-id="8a62c-580">function</span></span>| <span data-ttu-id="8a62c-581">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-581">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-582">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a62c-583">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8a62c-584">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a62c-585">エラー</span><span class="sxs-lookup"><span data-stu-id="8a62c-585">Errors</span></span>

| <span data-ttu-id="8a62c-586">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8a62c-586">Error code</span></span> | <span data-ttu-id="8a62c-587">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8a62c-588">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-589">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-589">Requirements</span></span>

|<span data-ttu-id="8a62c-590">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-590">Requirement</span></span>| <span data-ttu-id="8a62c-591">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-593">1.1</span><span class="sxs-lookup"><span data-stu-id="8a62c-593">1.1</span></span>|
|[<span data-ttu-id="8a62c-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a62c-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-597">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-598">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-598">Example</span></span>

<span data-ttu-id="8a62c-599">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="8a62c-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8a62c-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="8a62c-601">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-602">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-602">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8a62c-603">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-603">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8a62c-604">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="8a62c-605">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-605">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="8a62c-606">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-606">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="8a62c-607">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-607">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-608">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-608">Parameters</span></span>

|<span data-ttu-id="8a62c-609">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-609">Name</span></span>| <span data-ttu-id="8a62c-610">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-610">Type</span></span>| <span data-ttu-id="8a62c-611">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8a62c-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8a62c-612">String &#124; Object</span></span>| |<span data-ttu-id="8a62c-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8a62c-615">**または**</span><span class="sxs-lookup"><span data-stu-id="8a62c-615">**OR**</span></span><br/><span data-ttu-id="8a62c-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8a62c-618">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-618">String</span></span> | <span data-ttu-id="8a62c-619">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-619">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8a62c-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8a62c-623">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-623">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-624">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8a62c-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8a62c-625">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-625">String</span></span> | | <span data-ttu-id="8a62c-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8a62c-628">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-628">String</span></span> | | <span data-ttu-id="8a62c-629">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8a62c-630">文字列</span><span class="sxs-lookup"><span data-stu-id="8a62c-630">String</span></span> | | <span data-ttu-id="8a62c-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8a62c-633">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-633">String</span></span> | | <span data-ttu-id="8a62c-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8a62c-637">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-637">function</span></span> | <span data-ttu-id="8a62c-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-638">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-639">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-640">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-640">Requirements</span></span>

|<span data-ttu-id="8a62c-641">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-641">Requirement</span></span>| <span data-ttu-id="8a62c-642">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-643">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-644">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-644">1.0</span></span>|
|[<span data-ttu-id="8a62c-645">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-646">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-647">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-648">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a62c-649">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-649">Examples</span></span>

<span data-ttu-id="8a62c-650">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8a62c-651">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8a62c-652">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8a62c-653">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8a62c-654">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8a62c-655">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="8a62c-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="8a62c-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="8a62c-657">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-658">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-658">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8a62c-659">Web 上の Outlook では、返信フォームは、3列表示のポップアップフォームとして表示され、2列または1列表示のポップアップフォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-659">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8a62c-660">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="8a62c-661">`formData.attachments`パラメーターで添付ファイルが指定されている場合、web 上の Outlook およびデスクトップクライアントは、すべての添付ファイルをダウンロードして、返信フォームに添付しようとします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-661">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="8a62c-662">添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-662">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="8a62c-663">表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-663">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-664">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-664">Parameters</span></span>

|<span data-ttu-id="8a62c-665">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-665">Name</span></span>| <span data-ttu-id="8a62c-666">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-666">Type</span></span>| <span data-ttu-id="8a62c-667">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8a62c-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="8a62c-668">String &#124; Object</span></span>| | <span data-ttu-id="8a62c-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8a62c-671">**または**</span><span class="sxs-lookup"><span data-stu-id="8a62c-671">**OR**</span></span><br/><span data-ttu-id="8a62c-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8a62c-674">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-674">String</span></span> | <span data-ttu-id="8a62c-675">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-675">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="8a62c-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="8a62c-679">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-679">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-680">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="8a62c-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="8a62c-681">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-681">String</span></span> | | <span data-ttu-id="8a62c-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="8a62c-684">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-684">String</span></span> | | <span data-ttu-id="8a62c-685">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="8a62c-686">文字列</span><span class="sxs-lookup"><span data-stu-id="8a62c-686">String</span></span> | | <span data-ttu-id="8a62c-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="8a62c-689">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-689">String</span></span> | | <span data-ttu-id="8a62c-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="8a62c-693">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-693">function</span></span> | <span data-ttu-id="8a62c-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-694">&lt;optional&gt;</span></span> | <span data-ttu-id="8a62c-695">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-696">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-696">Requirements</span></span>

|<span data-ttu-id="8a62c-697">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-697">Requirement</span></span>| <span data-ttu-id="8a62c-698">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-699">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-700">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-700">1.0</span></span>|
|[<span data-ttu-id="8a62c-701">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-702">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-703">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-704">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8a62c-705">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-705">Examples</span></span>

<span data-ttu-id="8a62c-706">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8a62c-707">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8a62c-708">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8a62c-709">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="8a62c-710">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="8a62c-711">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="8a62c-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8a62c-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="8a62c-713">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-714">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-714">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-715">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-715">Requirements</span></span>

|<span data-ttu-id="8a62c-716">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-716">Requirement</span></span>| <span data-ttu-id="8a62c-717">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-718">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-719">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-719">1.0</span></span>|
|[<span data-ttu-id="8a62c-720">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-721">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-722">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-723">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-724">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-724">Returns:</span></span>

<span data-ttu-id="8a62c-725">型:[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8a62c-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8a62c-726">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-726">Example</span></span>

<span data-ttu-id="8a62c-727">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="8a62c-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8a62c-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8a62c-729">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-730">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-730">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-731">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-731">Parameters</span></span>

|<span data-ttu-id="8a62c-732">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-732">Name</span></span>| <span data-ttu-id="8a62c-733">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-733">Type</span></span>| <span data-ttu-id="8a62c-734">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8a62c-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8a62c-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="8a62c-736">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="8a62c-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a62c-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="8a62c-737">Requirements</span></span>

|<span data-ttu-id="8a62c-738">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-738">Requirement</span></span>| <span data-ttu-id="8a62c-739">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-740">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-741">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-741">1.0</span></span>|
|[<span data-ttu-id="8a62c-742">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-743">制限あり</span><span class="sxs-lookup"><span data-stu-id="8a62c-743">Restricted</span></span>|
|[<span data-ttu-id="8a62c-744">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-745">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-746">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-746">Returns:</span></span>

<span data-ttu-id="8a62c-747">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8a62c-748">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8a62c-749">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8a62c-750">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8a62c-751">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="8a62c-751">Value of `entityType`</span></span> | <span data-ttu-id="8a62c-752">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="8a62c-752">Type of objects in returned array</span></span> | <span data-ttu-id="8a62c-753">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8a62c-754">文字列</span><span class="sxs-lookup"><span data-stu-id="8a62c-754">String</span></span> | <span data-ttu-id="8a62c-755">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8a62c-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8a62c-756">連絡先</span><span class="sxs-lookup"><span data-stu-id="8a62c-756">Contact</span></span> | <span data-ttu-id="8a62c-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a62c-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8a62c-758">文字列</span><span class="sxs-lookup"><span data-stu-id="8a62c-758">String</span></span> | <span data-ttu-id="8a62c-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a62c-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8a62c-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8a62c-760">MeetingSuggestion</span></span> | <span data-ttu-id="8a62c-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a62c-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8a62c-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8a62c-762">PhoneNumber</span></span> | <span data-ttu-id="8a62c-763">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8a62c-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8a62c-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8a62c-764">TaskSuggestion</span></span> | <span data-ttu-id="8a62c-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8a62c-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8a62c-766">文字列</span><span class="sxs-lookup"><span data-stu-id="8a62c-766">String</span></span> | <span data-ttu-id="8a62c-767">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="8a62c-767">**Restricted**</span></span> |

<span data-ttu-id="8a62c-768">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8a62c-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="8a62c-769">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-769">Example</span></span>

<span data-ttu-id="8a62c-770">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="8a62c-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8a62c-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8a62c-772">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-773">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-773">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8a62c-774">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-775">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-775">Parameters</span></span>

|<span data-ttu-id="8a62c-776">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-776">Name</span></span>| <span data-ttu-id="8a62c-777">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-777">Type</span></span>| <span data-ttu-id="8a62c-778">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8a62c-779">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-779">String</span></span>|<span data-ttu-id="8a62c-780">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8a62c-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a62c-781">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-781">Requirements</span></span>

|<span data-ttu-id="8a62c-782">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-782">Requirement</span></span>| <span data-ttu-id="8a62c-783">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-784">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-785">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-785">1.0</span></span>|
|[<span data-ttu-id="8a62c-786">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-787">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-788">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-789">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-790">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-790">Returns:</span></span>

<span data-ttu-id="8a62c-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="8a62c-793">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8a62c-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="8a62c-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8a62c-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8a62c-795">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-796">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-796">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8a62c-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8a62c-800">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8a62c-801">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="8a62c-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a62c-804">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-804">Requirements</span></span>

|<span data-ttu-id="8a62c-805">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-805">Requirement</span></span>| <span data-ttu-id="8a62c-806">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-807">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-808">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-808">1.0</span></span>|
|[<span data-ttu-id="8a62c-809">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-810">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-811">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-812">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-813">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-813">Returns:</span></span>

<span data-ttu-id="8a62c-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8a62c-816">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a62c-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a62c-817">Object</span><span class="sxs-lookup"><span data-stu-id="8a62c-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a62c-818">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-818">Example</span></span>

<span data-ttu-id="8a62c-819">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="8a62c-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8a62c-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="8a62c-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8a62c-821">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8a62c-822">このメソッドは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-822">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8a62c-823">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8a62c-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-826">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-826">Parameters</span></span>

|<span data-ttu-id="8a62c-827">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-827">Name</span></span>| <span data-ttu-id="8a62c-828">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-828">Type</span></span>| <span data-ttu-id="8a62c-829">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8a62c-830">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-830">String</span></span>|<span data-ttu-id="8a62c-831">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="8a62c-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a62c-832">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-832">Requirements</span></span>

|<span data-ttu-id="8a62c-833">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-833">Requirement</span></span>| <span data-ttu-id="8a62c-834">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-835">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-836">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-836">1.0</span></span>|
|[<span data-ttu-id="8a62c-837">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-838">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-839">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-840">読み取り</span><span class="sxs-lookup"><span data-stu-id="8a62c-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-841">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-841">Returns:</span></span>

<span data-ttu-id="8a62c-842">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="8a62c-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8a62c-843">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a62c-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a62c-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8a62c-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a62c-845">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="8a62c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="8a62c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="8a62c-847">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="8a62c-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-850">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-850">Parameters</span></span>

|<span data-ttu-id="8a62c-851">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-851">Name</span></span>| <span data-ttu-id="8a62c-852">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-852">Type</span></span>| <span data-ttu-id="8a62c-853">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-853">Attributes</span></span>| <span data-ttu-id="8a62c-854">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="8a62c-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8a62c-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="8a62c-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="8a62c-859">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-859">Object</span></span>| <span data-ttu-id="8a62c-860">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-860">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-861">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8a62c-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8a62c-862">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-862">Object</span></span>| <span data-ttu-id="8a62c-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-863">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-864">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8a62c-865">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-865">function</span></span>||<span data-ttu-id="8a62c-866">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a62c-867">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="8a62c-868">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="8a62c-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a62c-869">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-869">Requirements</span></span>

|<span data-ttu-id="8a62c-870">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-870">Requirement</span></span>| <span data-ttu-id="8a62c-871">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-872">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-873">1.2</span><span class="sxs-lookup"><span data-stu-id="8a62c-873">1.2</span></span>|
|[<span data-ttu-id="8a62c-874">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a62c-876">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-877">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="8a62c-878">戻り値:</span><span class="sxs-lookup"><span data-stu-id="8a62c-878">Returns:</span></span>

<span data-ttu-id="8a62c-879">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="8a62c-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="8a62c-880">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8a62c-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8a62c-881">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8a62c-882">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-882">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8a62c-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8a62c-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8a62c-884">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8a62c-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-888">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-888">Parameters</span></span>

|<span data-ttu-id="8a62c-889">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-889">Name</span></span>| <span data-ttu-id="8a62c-890">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-890">Type</span></span>| <span data-ttu-id="8a62c-891">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-891">Attributes</span></span>| <span data-ttu-id="8a62c-892">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8a62c-893">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-893">function</span></span>||<span data-ttu-id="8a62c-894">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8a62c-895">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8a62c-896">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8a62c-897">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-897">Object</span></span>| <span data-ttu-id="8a62c-898">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-898">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-899">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8a62c-900">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a62c-901">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-901">Requirements</span></span>

|<span data-ttu-id="8a62c-902">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-902">Requirement</span></span>| <span data-ttu-id="8a62c-903">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-904">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-905">1.0</span><span class="sxs-lookup"><span data-stu-id="8a62c-905">1.0</span></span>|
|[<span data-ttu-id="8a62c-906">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-907">ReadItem</span></span>|
|[<span data-ttu-id="8a62c-908">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-909">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8a62c-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-910">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-910">Example</span></span>

<span data-ttu-id="8a62c-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8a62c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8a62c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8a62c-915">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8a62c-916">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-916">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="8a62c-917">ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="8a62c-917">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="8a62c-918">Outlook on the web およびモバイルデバイスでは、添付ファイル識別子は同じセッション内でのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="8a62c-918">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="8a62c-919">ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-919">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-920">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-920">Parameters</span></span>

|<span data-ttu-id="8a62c-921">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-921">Name</span></span>| <span data-ttu-id="8a62c-922">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-922">Type</span></span>| <span data-ttu-id="8a62c-923">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-923">Attributes</span></span>| <span data-ttu-id="8a62c-924">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8a62c-925">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-925">String</span></span>||<span data-ttu-id="8a62c-926">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="8a62c-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8a62c-927">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-927">Object</span></span>| <span data-ttu-id="8a62c-928">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-928">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-929">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8a62c-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8a62c-930">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-930">Object</span></span>| <span data-ttu-id="8a62c-931">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-931">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-932">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8a62c-933">関数</span><span class="sxs-lookup"><span data-stu-id="8a62c-933">function</span></span>| <span data-ttu-id="8a62c-934">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-934">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-935">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8a62c-936">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8a62c-937">エラー</span><span class="sxs-lookup"><span data-stu-id="8a62c-937">Errors</span></span>

| <span data-ttu-id="8a62c-938">エラー コード</span><span class="sxs-lookup"><span data-stu-id="8a62c-938">Error code</span></span> | <span data-ttu-id="8a62c-939">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8a62c-940">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="8a62c-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-941">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-941">Requirements</span></span>

|<span data-ttu-id="8a62c-942">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-942">Requirement</span></span>| <span data-ttu-id="8a62c-943">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-944">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-945">1.1</span><span class="sxs-lookup"><span data-stu-id="8a62c-945">1.1</span></span>|
|[<span data-ttu-id="8a62c-946">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a62c-948">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-949">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-950">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-950">Example</span></span>

<span data-ttu-id="8a62c-951">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="8a62c-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="8a62c-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="8a62c-953">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="8a62c-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="8a62c-p166">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8a62c-957">パラメーター</span><span class="sxs-lookup"><span data-stu-id="8a62c-957">Parameters</span></span>

|<span data-ttu-id="8a62c-958">名前</span><span class="sxs-lookup"><span data-stu-id="8a62c-958">Name</span></span>| <span data-ttu-id="8a62c-959">型</span><span class="sxs-lookup"><span data-stu-id="8a62c-959">Type</span></span>| <span data-ttu-id="8a62c-960">属性</span><span class="sxs-lookup"><span data-stu-id="8a62c-960">Attributes</span></span>| <span data-ttu-id="8a62c-961">説明</span><span class="sxs-lookup"><span data-stu-id="8a62c-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8a62c-962">String</span><span class="sxs-lookup"><span data-stu-id="8a62c-962">String</span></span>||<span data-ttu-id="8a62c-p167">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="8a62c-966">Object</span><span class="sxs-lookup"><span data-stu-id="8a62c-966">Object</span></span>| <span data-ttu-id="8a62c-967">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-967">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-968">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="8a62c-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8a62c-969">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8a62c-969">Object</span></span>| <span data-ttu-id="8a62c-970">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-970">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-971">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="8a62c-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="8a62c-972">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="8a62c-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="8a62c-973">&lt;optional&gt;</span></span>|<span data-ttu-id="8a62c-974">の`text`場合、現在のスタイルが Outlook on the web およびデスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-974">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="8a62c-975">フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-975">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="8a62c-976">フィールド`html`が HTML をサポートする場合 (件名は含まれません)、現在のスタイルが outlook on the web で適用され、既定のスタイルが outlook デスクトップクライアントで適用されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-976">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="8a62c-977">フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-977">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="8a62c-978">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="8a62c-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="8a62c-979">function</span><span class="sxs-lookup"><span data-stu-id="8a62c-979">function</span></span>||<span data-ttu-id="8a62c-980">メソッドが完了すると、`callback` パラメータに渡された関数が、単一パラメータ `asyncResult` で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="8a62c-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a62c-981">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-981">Requirements</span></span>

|<span data-ttu-id="8a62c-982">要件</span><span class="sxs-lookup"><span data-stu-id="8a62c-982">Requirement</span></span>| <span data-ttu-id="8a62c-983">値</span><span class="sxs-lookup"><span data-stu-id="8a62c-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a62c-984">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8a62c-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a62c-985">1.2</span><span class="sxs-lookup"><span data-stu-id="8a62c-985">1.2</span></span>|
|[<span data-ttu-id="8a62c-986">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8a62c-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a62c-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8a62c-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="8a62c-988">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8a62c-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a62c-989">作成</span><span class="sxs-lookup"><span data-stu-id="8a62c-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8a62c-990">例</span><span class="sxs-lookup"><span data-stu-id="8a62c-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
