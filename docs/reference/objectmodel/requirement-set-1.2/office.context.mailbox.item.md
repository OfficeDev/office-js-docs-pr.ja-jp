---
title: Office. メールボックス-要件セット1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: c0b956cac0410ef7d8e8e0d59a69e221e29c540a
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068043"
---
# <a name="item"></a><span data-ttu-id="4c2e0-102">item</span><span class="sxs-lookup"><span data-stu-id="4c2e0-102">item</span></span>

### <span data-ttu-id="4c2e0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="4c2e0-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-107">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-107">Requirements</span></span>

|<span data-ttu-id="4c2e0-108">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-108">Requirement</span></span>| <span data-ttu-id="4c2e0-109">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-111">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-111">1.0</span></span>|
|[<span data-ttu-id="4c2e0-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="4c2e0-113">Restricted</span></span>|
|[<span data-ttu-id="4c2e0-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-115">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="4c2e0-116">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-116">Example</span></span>

<span data-ttu-id="4c2e0-117">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4c2e0-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="4c2e0-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="4c2e0-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4c2e0-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="4c2e0-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-122">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4c2e0-123">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-124">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-124">Type</span></span>

*   <span data-ttu-id="4c2e0-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4c2e0-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-126">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-126">Requirements</span></span>

|<span data-ttu-id="4c2e0-127">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-127">Requirement</span></span>| <span data-ttu-id="4c2e0-128">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-129">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-130">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-130">1.0</span></span>|
|[<span data-ttu-id="4c2e0-131">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-132">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-134">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-135">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-135">Example</span></span>

<span data-ttu-id="4c2e0-136">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4c2e0-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4c2e0-138">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4c2e0-139">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-140">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-140">Type</span></span>

*   [<span data-ttu-id="4c2e0-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="4c2e0-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-142">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-142">Requirements</span></span>

|<span data-ttu-id="4c2e0-143">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-143">Requirement</span></span>| <span data-ttu-id="4c2e0-144">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-146">1.1</span><span class="sxs-lookup"><span data-stu-id="4c2e0-146">1.1</span></span>|
|[<span data-ttu-id="4c2e0-147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-148">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-150">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-151">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="4c2e0-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="4c2e0-153">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-154">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-154">Type</span></span>

*   [<span data-ttu-id="4c2e0-155">Body</span><span class="sxs-lookup"><span data-stu-id="4c2e0-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-156">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-156">Requirements</span></span>

|<span data-ttu-id="4c2e0-157">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-157">Requirement</span></span>| <span data-ttu-id="4c2e0-158">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-160">1.1</span><span class="sxs-lookup"><span data-stu-id="4c2e0-160">1.1</span></span>|
|[<span data-ttu-id="4c2e0-161">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-162">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-164">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-165">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-165">Example</span></span>

<span data-ttu-id="4c2e0-166">この例では、メッセージの本文をプレーンテキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4c2e0-167">次の例は、コールバック関数に渡される result パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4c2e0-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4c2e0-169">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4c2e0-170">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-171">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-171">Read mode</span></span>

<span data-ttu-id="4c2e0-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-174">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-174">Compose mode</span></span>

<span data-ttu-id="4c2e0-175">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-176">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-176">Type</span></span>

*   <span data-ttu-id="4c2e0-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-178">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-178">Requirements</span></span>

|<span data-ttu-id="4c2e0-179">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-179">Requirement</span></span>| <span data-ttu-id="4c2e0-180">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-182">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-182">1.0</span></span>|
|[<span data-ttu-id="4c2e0-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-184">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-186">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4c2e0-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="4c2e0-188">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4c2e0-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4c2e0-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-193">Type</span><span class="sxs-lookup"><span data-stu-id="4c2e0-193">Type</span></span>

*   <span data-ttu-id="4c2e0-194">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-195">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-195">Requirements</span></span>

|<span data-ttu-id="4c2e0-196">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-196">Requirement</span></span>| <span data-ttu-id="4c2e0-197">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-199">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-199">1.0</span></span>|
|[<span data-ttu-id="4c2e0-200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-200">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-201">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-202">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-203">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-204">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="4c2e0-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4c2e0-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="4c2e0-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-208">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-208">Type</span></span>

*   <span data-ttu-id="4c2e0-209">日付</span><span class="sxs-lookup"><span data-stu-id="4c2e0-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-210">Requirements</span></span>

|<span data-ttu-id="4c2e0-211">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-211">Requirement</span></span>| <span data-ttu-id="4c2e0-212">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-214">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-214">1.0</span></span>|
|[<span data-ttu-id="4c2e0-215">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-215">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-216">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-217">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-218">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-219">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4c2e0-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4c2e0-220">dateTimeModified :Date</span></span>

<span data-ttu-id="4c2e0-p111">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-223">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-224">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-224">Type</span></span>

*   <span data-ttu-id="4c2e0-225">日付</span><span class="sxs-lookup"><span data-stu-id="4c2e0-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-226">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-226">Requirements</span></span>

|<span data-ttu-id="4c2e0-227">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-227">Requirement</span></span>| <span data-ttu-id="4c2e0-228">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-229">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-230">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-230">1.0</span></span>|
|[<span data-ttu-id="4c2e0-231">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-231">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-232">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-233">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-233">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-234">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-235">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="4c2e0-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="4c2e0-237">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4c2e0-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-240">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-240">Read mode</span></span>

<span data-ttu-id="4c2e0-241">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-242">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-242">Compose mode</span></span>

<span data-ttu-id="4c2e0-243">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4c2e0-244">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4c2e0-245">次の例では、 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) `Time`オブジェクトのメソッドを使用して予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4c2e0-246">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-246">Type</span></span>

*   <span data-ttu-id="4c2e0-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-248">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-248">Requirements</span></span>

|<span data-ttu-id="4c2e0-249">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-249">Requirement</span></span>| <span data-ttu-id="4c2e0-250">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-252">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-252">1.0</span></span>|
|[<span data-ttu-id="4c2e0-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-253">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-254">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-255">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-256">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4c2e0-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4c2e0-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4c2e0-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-262">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-263">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-263">Type</span></span>

*   [<span data-ttu-id="4c2e0-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4c2e0-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-265">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-265">Requirements</span></span>

|<span data-ttu-id="4c2e0-266">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-266">Requirement</span></span>| <span data-ttu-id="4c2e0-267">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-269">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-269">1.0</span></span>|
|[<span data-ttu-id="4c2e0-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-271">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-273">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-274">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="4c2e0-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-275">internetMessageId :String</span></span>

<span data-ttu-id="4c2e0-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-278">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-278">Type</span></span>

*   <span data-ttu-id="4c2e0-279">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-280">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-280">Requirements</span></span>

|<span data-ttu-id="4c2e0-281">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-281">Requirement</span></span>| <span data-ttu-id="4c2e0-282">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-283">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-284">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-284">1.0</span></span>|
|[<span data-ttu-id="4c2e0-285">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-285">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-286">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-287">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-288">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-289">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4c2e0-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-290">itemClass :String</span></span>

<span data-ttu-id="4c2e0-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4c2e0-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4c2e0-295">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-295">Type</span></span> | <span data-ttu-id="4c2e0-296">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-296">Description</span></span> | <span data-ttu-id="4c2e0-297">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="4c2e0-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4c2e0-298">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="4c2e0-298">Appointment items</span></span> | <span data-ttu-id="4c2e0-299">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="4c2e0-300">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="4c2e0-300">Message items</span></span> | <span data-ttu-id="4c2e0-301">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4c2e0-302">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-303">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-303">Type</span></span>

*   <span data-ttu-id="4c2e0-304">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-305">Requirements</span></span>

|<span data-ttu-id="4c2e0-306">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-306">Requirement</span></span>| <span data-ttu-id="4c2e0-307">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-309">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-309">1.0</span></span>|
|[<span data-ttu-id="4c2e0-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-311">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-313">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-314">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4c2e0-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-315">(nullable) itemId :String</span></span>

<span data-ttu-id="4c2e0-p118">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-318">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4c2e0-319">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4c2e0-320">この値を使用して REST API を呼び出す前に、要件セット 1.3 から使用できる `Office.context.mailbox.convertToRestId` を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="4c2e0-321">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-322">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-322">Type</span></span>

*   <span data-ttu-id="4c2e0-323">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-324">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-324">Requirements</span></span>

|<span data-ttu-id="4c2e0-325">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-325">Requirement</span></span>| <span data-ttu-id="4c2e0-326">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-328">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-328">1.0</span></span>|
|[<span data-ttu-id="4c2e0-329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-330">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-332">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-333">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-333">Example</span></span>

<span data-ttu-id="4c2e0-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="4c2e0-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4c2e0-337">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4c2e0-338">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-339">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-339">Type</span></span>

*   [<span data-ttu-id="4c2e0-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4c2e0-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-341">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-341">Requirements</span></span>

|<span data-ttu-id="4c2e0-342">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-342">Requirement</span></span>| <span data-ttu-id="4c2e0-343">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-344">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-345">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-345">1.0</span></span>|
|[<span data-ttu-id="4c2e0-346">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-347">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-348">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-349">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-350">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="4c2e0-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="4c2e0-352">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-353">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-353">Read mode</span></span>

<span data-ttu-id="4c2e0-354">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-355">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-355">Compose mode</span></span>

<span data-ttu-id="4c2e0-356">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-357">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-357">Type</span></span>

*   <span data-ttu-id="4c2e0-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-359">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-359">Requirements</span></span>

|<span data-ttu-id="4c2e0-360">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-360">Requirement</span></span>| <span data-ttu-id="4c2e0-361">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-363">1.0</span></span>|
|[<span data-ttu-id="4c2e0-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-365">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-367">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4c2e0-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-368">normalizedSubject :String</span></span>

<span data-ttu-id="4c2e0-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4c2e0-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-373">Type</span><span class="sxs-lookup"><span data-stu-id="4c2e0-373">Type</span></span>

*   <span data-ttu-id="4c2e0-374">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-375">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-375">Requirements</span></span>

|<span data-ttu-id="4c2e0-376">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-376">Requirement</span></span>| <span data-ttu-id="4c2e0-377">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-378">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-379">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-379">1.0</span></span>|
|[<span data-ttu-id="4c2e0-380">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-381">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-382">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-383">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-384">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4c2e0-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4c2e0-386">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4c2e0-387">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-388">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-388">Read mode</span></span>

<span data-ttu-id="4c2e0-389">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-390">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-390">Compose mode</span></span>

<span data-ttu-id="4c2e0-391">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-392">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-392">Type</span></span>

*   <span data-ttu-id="4c2e0-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-394">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-394">Requirements</span></span>

|<span data-ttu-id="4c2e0-395">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-395">Requirement</span></span>| <span data-ttu-id="4c2e0-396">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-397">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-398">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-398">1.0</span></span>|
|[<span data-ttu-id="4c2e0-399">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-400">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-401">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-402">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4c2e0-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4c2e0-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-406">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-406">Type</span></span>

*   [<span data-ttu-id="4c2e0-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4c2e0-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-408">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-408">Requirements</span></span>

|<span data-ttu-id="4c2e0-409">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-409">Requirement</span></span>| <span data-ttu-id="4c2e0-410">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-412">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-412">1.0</span></span>|
|[<span data-ttu-id="4c2e0-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-414">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-417">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4c2e0-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4c2e0-419">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4c2e0-420">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-421">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-421">Read mode</span></span>

<span data-ttu-id="4c2e0-422">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-423">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-423">Compose mode</span></span>

<span data-ttu-id="4c2e0-424">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-425">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-425">Type</span></span>

*   <span data-ttu-id="4c2e0-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-427">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-427">Requirements</span></span>

|<span data-ttu-id="4c2e0-428">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-428">Requirement</span></span>| <span data-ttu-id="4c2e0-429">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-431">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-431">1.0</span></span>|
|[<span data-ttu-id="4c2e0-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-433">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-435">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4c2e0-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4c2e0-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4c2e0-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-441">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4c2e0-442">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-442">Type</span></span>

*   [<span data-ttu-id="4c2e0-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4c2e0-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4c2e0-444">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-444">Requirements</span></span>

|<span data-ttu-id="4c2e0-445">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-445">Requirement</span></span>| <span data-ttu-id="4c2e0-446">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-448">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-448">1.0</span></span>|
|[<span data-ttu-id="4c2e0-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-450">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-453">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="4c2e0-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="4c2e0-455">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4c2e0-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-458">Read mode</span></span>

<span data-ttu-id="4c2e0-459">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-460">Compose mode</span></span>

<span data-ttu-id="4c2e0-461">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4c2e0-462">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="4c2e0-463">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4c2e0-464">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-464">Type</span></span>

*   <span data-ttu-id="4c2e0-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-466">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-466">Requirements</span></span>

|<span data-ttu-id="4c2e0-467">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-467">Requirement</span></span>| <span data-ttu-id="4c2e0-468">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-470">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-470">1.0</span></span>|
|[<span data-ttu-id="4c2e0-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-472">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-474">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="4c2e0-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="4c2e0-476">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4c2e0-477">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-478">Read mode</span></span>

<span data-ttu-id="4c2e0-p130">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-481">Compose mode</span></span>

<span data-ttu-id="4c2e0-482">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-483">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-483">Type</span></span>

*   <span data-ttu-id="4c2e0-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-485">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-485">Requirements</span></span>

|<span data-ttu-id="4c2e0-486">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-486">Requirement</span></span>| <span data-ttu-id="4c2e0-487">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-489">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-489">1.0</span></span>|
|[<span data-ttu-id="4c2e0-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-491">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-493">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4c2e0-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4c2e0-495">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4c2e0-496">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4c2e0-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-497">Read mode</span></span>

<span data-ttu-id="4c2e0-p132">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4c2e0-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-500">Compose mode</span></span>

<span data-ttu-id="4c2e0-501">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4c2e0-502">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-502">Type</span></span>

*   <span data-ttu-id="4c2e0-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-504">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-504">Requirements</span></span>

|<span data-ttu-id="4c2e0-505">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-505">Requirement</span></span>| <span data-ttu-id="4c2e0-506">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-508">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-508">1.0</span></span>|
|[<span data-ttu-id="4c2e0-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-510">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-512">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4c2e0-513">メソッド</span><span class="sxs-lookup"><span data-stu-id="4c2e0-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4c2e0-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4c2e0-515">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4c2e0-516">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4c2e0-517">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-518">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-518">Parameters</span></span>

|<span data-ttu-id="4c2e0-519">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-519">Name</span></span>| <span data-ttu-id="4c2e0-520">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-520">Type</span></span>| <span data-ttu-id="4c2e0-521">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-521">Attributes</span></span>| <span data-ttu-id="4c2e0-522">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4c2e0-523">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-523">String</span></span>||<span data-ttu-id="4c2e0-p133">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4c2e0-526">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-526">String</span></span>||<span data-ttu-id="4c2e0-p134">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4c2e0-529">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-529">Object</span></span>| <span data-ttu-id="4c2e0-530">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-530">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-531">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4c2e0-532">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-532">Object</span></span>| <span data-ttu-id="4c2e0-533">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-533">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-534">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4c2e0-535">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-535">function</span></span>| <span data-ttu-id="4c2e0-536">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-536">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-537">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4c2e0-538">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4c2e0-539">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4c2e0-540">エラー</span><span class="sxs-lookup"><span data-stu-id="4c2e0-540">Errors</span></span>

| <span data-ttu-id="4c2e0-541">エラー コード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-541">Error code</span></span> | <span data-ttu-id="4c2e0-542">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4c2e0-543">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4c2e0-544">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4c2e0-545">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-546">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-546">Requirements</span></span>

|<span data-ttu-id="4c2e0-547">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-547">Requirement</span></span>| <span data-ttu-id="4c2e0-548">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-550">1.1</span><span class="sxs-lookup"><span data-stu-id="4c2e0-550">1.1</span></span>|
|[<span data-ttu-id="4c2e0-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="4c2e0-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-554">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-555">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4c2e0-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4c2e0-557">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4c2e0-p135">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4c2e0-561">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4c2e0-562">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-563">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-563">Parameters</span></span>

|<span data-ttu-id="4c2e0-564">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-564">Name</span></span>| <span data-ttu-id="4c2e0-565">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-565">Type</span></span>| <span data-ttu-id="4c2e0-566">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-566">Attributes</span></span>| <span data-ttu-id="4c2e0-567">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4c2e0-568">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-568">String</span></span>||<span data-ttu-id="4c2e0-p136">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4c2e0-571">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-571">String</span></span>||<span data-ttu-id="4c2e0-572">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-572">The subject of the item to be attached.</span></span> <span data-ttu-id="4c2e0-573">最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4c2e0-574">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-574">Object</span></span>| <span data-ttu-id="4c2e0-575">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-575">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-576">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4c2e0-577">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-577">Object</span></span>| <span data-ttu-id="4c2e0-578">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-578">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-579">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4c2e0-580">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-580">function</span></span>| <span data-ttu-id="4c2e0-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-581">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-582">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4c2e0-583">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4c2e0-584">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4c2e0-585">エラー</span><span class="sxs-lookup"><span data-stu-id="4c2e0-585">Errors</span></span>

| <span data-ttu-id="4c2e0-586">エラー コード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-586">Error code</span></span> | <span data-ttu-id="4c2e0-587">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4c2e0-588">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-589">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-589">Requirements</span></span>

|<span data-ttu-id="4c2e0-590">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-590">Requirement</span></span>| <span data-ttu-id="4c2e0-591">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-593">1.1</span><span class="sxs-lookup"><span data-stu-id="4c2e0-593">1.1</span></span>|
|[<span data-ttu-id="4c2e0-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="4c2e0-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-597">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-598">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-598">Example</span></span>

<span data-ttu-id="4c2e0-599">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4c2e0-600">displayReplyAllForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4c2e0-601">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-602">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c2e0-603">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4c2e0-604">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4c2e0-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-608">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-608">Parameters</span></span>

|<span data-ttu-id="4c2e0-609">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-609">Name</span></span>| <span data-ttu-id="4c2e0-610">種類</span><span class="sxs-lookup"><span data-stu-id="4c2e0-610">Type</span></span>| <span data-ttu-id="4c2e0-611">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4c2e0-612">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-612">String &#124; Object</span></span>| |<span data-ttu-id="4c2e0-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4c2e0-615">**または**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-615">**OR**</span></span><br/><span data-ttu-id="4c2e0-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4c2e0-618">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-618">String</span></span> | <span data-ttu-id="4c2e0-619">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-619">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4c2e0-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4c2e0-623">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-623">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-624">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4c2e0-625">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-625">String</span></span> | | <span data-ttu-id="4c2e0-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4c2e0-628">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-628">String</span></span> | | <span data-ttu-id="4c2e0-629">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4c2e0-630">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-630">String</span></span> | | <span data-ttu-id="4c2e0-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4c2e0-633">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-633">String</span></span> | | <span data-ttu-id="4c2e0-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4c2e0-637">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-637">function</span></span> | <span data-ttu-id="4c2e0-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-638">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-639">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-640">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-640">Requirements</span></span>

|<span data-ttu-id="4c2e0-641">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-641">Requirement</span></span>| <span data-ttu-id="4c2e0-642">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-643">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-644">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-644">1.0</span></span>|
|[<span data-ttu-id="4c2e0-645">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-645">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-646">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-647">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-647">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-648">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4c2e0-649">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-649">Examples</span></span>

<span data-ttu-id="4c2e0-650">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4c2e0-651">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4c2e0-652">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4c2e0-653">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4c2e0-654">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4c2e0-655">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4c2e0-656">displayReplyForm (formdata, [callback])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4c2e0-657">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-658">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c2e0-659">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4c2e0-660">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4c2e0-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-664">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-664">Parameters</span></span>

|<span data-ttu-id="4c2e0-665">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-665">Name</span></span>| <span data-ttu-id="4c2e0-666">種類</span><span class="sxs-lookup"><span data-stu-id="4c2e0-666">Type</span></span>| <span data-ttu-id="4c2e0-667">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4c2e0-668">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-668">String &#124; Object</span></span>| | <span data-ttu-id="4c2e0-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4c2e0-671">**または**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-671">**OR**</span></span><br/><span data-ttu-id="4c2e0-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4c2e0-674">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-674">String</span></span> | <span data-ttu-id="4c2e0-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-675">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4c2e0-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4c2e0-679">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-679">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-680">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4c2e0-681">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-681">String</span></span> | | <span data-ttu-id="4c2e0-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4c2e0-684">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-684">String</span></span> | | <span data-ttu-id="4c2e0-685">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4c2e0-686">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-686">String</span></span> | | <span data-ttu-id="4c2e0-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4c2e0-689">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-689">String</span></span> | | <span data-ttu-id="4c2e0-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4c2e0-693">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-693">function</span></span> | <span data-ttu-id="4c2e0-694">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-694">&lt;optional&gt;</span></span> | <span data-ttu-id="4c2e0-695">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-696">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-696">Requirements</span></span>

|<span data-ttu-id="4c2e0-697">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-697">Requirement</span></span>| <span data-ttu-id="4c2e0-698">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-699">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-700">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-700">1.0</span></span>|
|[<span data-ttu-id="4c2e0-701">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-701">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-702">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-703">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-703">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-704">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4c2e0-705">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-705">Examples</span></span>

<span data-ttu-id="4c2e0-706">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4c2e0-707">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4c2e0-708">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4c2e0-709">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4c2e0-710">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4c2e0-711">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="4c2e0-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="4c2e0-713">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-714">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-715">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-715">Requirements</span></span>

|<span data-ttu-id="4c2e0-716">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-716">Requirement</span></span>| <span data-ttu-id="4c2e0-717">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-718">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-719">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-719">1.0</span></span>|
|[<span data-ttu-id="4c2e0-720">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-720">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-721">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-722">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-722">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-723">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-724">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-724">Returns:</span></span>

<span data-ttu-id="4c2e0-725">型:[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4c2e0-726">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-726">Example</span></span>

<span data-ttu-id="4c2e0-727">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="4c2e0-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4c2e0-729">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-730">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-731">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-731">Parameters</span></span>

|<span data-ttu-id="4c2e0-732">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-732">Name</span></span>| <span data-ttu-id="4c2e0-733">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-733">Type</span></span>| <span data-ttu-id="4c2e0-734">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4c2e0-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4c2e0-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="4c2e0-736">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c2e0-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="4c2e0-737">Requirements</span></span>

|<span data-ttu-id="4c2e0-738">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-738">Requirement</span></span>| <span data-ttu-id="4c2e0-739">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-740">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-741">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-741">1.0</span></span>|
|[<span data-ttu-id="4c2e0-742">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-742">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-743">制限あり</span><span class="sxs-lookup"><span data-stu-id="4c2e0-743">Restricted</span></span>|
|[<span data-ttu-id="4c2e0-744">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-744">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-745">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-746">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-746">Returns:</span></span>

<span data-ttu-id="4c2e0-747">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4c2e0-748">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4c2e0-749">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4c2e0-750">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4c2e0-751">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-751">Value of `entityType`</span></span> | <span data-ttu-id="4c2e0-752">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-752">Type of objects in returned array</span></span> | <span data-ttu-id="4c2e0-753">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4c2e0-754">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-754">String</span></span> | <span data-ttu-id="4c2e0-755">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4c2e0-756">連絡先</span><span class="sxs-lookup"><span data-stu-id="4c2e0-756">Contact</span></span> | <span data-ttu-id="4c2e0-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4c2e0-758">文字列</span><span class="sxs-lookup"><span data-stu-id="4c2e0-758">String</span></span> | <span data-ttu-id="4c2e0-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4c2e0-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4c2e0-760">MeetingSuggestion</span></span> | <span data-ttu-id="4c2e0-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4c2e0-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4c2e0-762">PhoneNumber</span></span> | <span data-ttu-id="4c2e0-763">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4c2e0-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4c2e0-764">TaskSuggestion</span></span> | <span data-ttu-id="4c2e0-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4c2e0-766">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-766">String</span></span> | <span data-ttu-id="4c2e0-767">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="4c2e0-767">**Restricted**</span></span> |

<span data-ttu-id="4c2e0-768">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4c2e0-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4c2e0-769">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-769">Example</span></span>

<span data-ttu-id="4c2e0-770">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="4c2e0-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4c2e0-772">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-773">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c2e0-774">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-775">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-775">Parameters</span></span>

|<span data-ttu-id="4c2e0-776">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-776">Name</span></span>| <span data-ttu-id="4c2e0-777">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-777">Type</span></span>| <span data-ttu-id="4c2e0-778">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4c2e0-779">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-779">String</span></span>|<span data-ttu-id="4c2e0-780">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c2e0-781">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-781">Requirements</span></span>

|<span data-ttu-id="4c2e0-782">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-782">Requirement</span></span>| <span data-ttu-id="4c2e0-783">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-784">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-785">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-785">1.0</span></span>|
|[<span data-ttu-id="4c2e0-786">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-786">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-787">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-788">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-788">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-789">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-790">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-790">Returns:</span></span>

<span data-ttu-id="4c2e0-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4c2e0-793">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4c2e0-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="4c2e0-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4c2e0-795">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-796">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c2e0-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4c2e0-800">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4c2e0-801">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="4c2e0-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c2e0-804">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-804">Requirements</span></span>

|<span data-ttu-id="4c2e0-805">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-805">Requirement</span></span>| <span data-ttu-id="4c2e0-806">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-807">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-808">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-808">1.0</span></span>|
|[<span data-ttu-id="4c2e0-809">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-810">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-811">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-812">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-813">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-813">Returns:</span></span>

<span data-ttu-id="4c2e0-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4c2e0-816">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="4c2e0-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4c2e0-817">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="4c2e0-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4c2e0-818">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-818">Example</span></span>

<span data-ttu-id="4c2e0-819">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="4c2e0-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4c2e0-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4c2e0-821">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4c2e0-822">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c2e0-823">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4c2e0-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-826">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-826">Parameters</span></span>

|<span data-ttu-id="4c2e0-827">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-827">Name</span></span>| <span data-ttu-id="4c2e0-828">種類</span><span class="sxs-lookup"><span data-stu-id="4c2e0-828">Type</span></span>| <span data-ttu-id="4c2e0-829">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4c2e0-830">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-830">String</span></span>|<span data-ttu-id="4c2e0-831">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c2e0-832">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-832">Requirements</span></span>

|<span data-ttu-id="4c2e0-833">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-833">Requirement</span></span>| <span data-ttu-id="4c2e0-834">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-835">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-836">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-836">1.0</span></span>|
|[<span data-ttu-id="4c2e0-837">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-838">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-839">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-840">読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-841">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-841">Returns:</span></span>

<span data-ttu-id="4c2e0-842">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4c2e0-843">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="4c2e0-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4c2e0-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4c2e0-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4c2e0-845">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4c2e0-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4c2e0-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4c2e0-847">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4c2e0-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-850">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-850">Parameters</span></span>

|<span data-ttu-id="4c2e0-851">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-851">Name</span></span>| <span data-ttu-id="4c2e0-852">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-852">Type</span></span>| <span data-ttu-id="4c2e0-853">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-853">Attributes</span></span>| <span data-ttu-id="4c2e0-854">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4c2e0-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4c2e0-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4c2e0-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4c2e0-859">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-859">Object</span></span>| <span data-ttu-id="4c2e0-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-860">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-861">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4c2e0-862">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-862">Object</span></span>| <span data-ttu-id="4c2e0-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-863">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-864">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4c2e0-865">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-865">function</span></span>||<span data-ttu-id="4c2e0-866">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4c2e0-867">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4c2e0-868">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c2e0-869">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-869">Requirements</span></span>

|<span data-ttu-id="4c2e0-870">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-870">Requirement</span></span>| <span data-ttu-id="4c2e0-871">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-872">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-873">1.2</span><span class="sxs-lookup"><span data-stu-id="4c2e0-873">1.2</span></span>|
|[<span data-ttu-id="4c2e0-874">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="4c2e0-876">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-877">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c2e0-878">戻り値:</span><span class="sxs-lookup"><span data-stu-id="4c2e0-878">Returns:</span></span>

<span data-ttu-id="4c2e0-879">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4c2e0-880">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="4c2e0-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4c2e0-881">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4c2e0-882">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-882">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4c2e0-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4c2e0-884">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4c2e0-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-888">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-888">Parameters</span></span>

|<span data-ttu-id="4c2e0-889">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-889">Name</span></span>| <span data-ttu-id="4c2e0-890">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-890">Type</span></span>| <span data-ttu-id="4c2e0-891">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-891">Attributes</span></span>| <span data-ttu-id="4c2e0-892">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4c2e0-893">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-893">function</span></span>||<span data-ttu-id="4c2e0-894">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4c2e0-895">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4c2e0-896">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4c2e0-897">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-897">Object</span></span>| <span data-ttu-id="4c2e0-898">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-898">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-899">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4c2e0-900">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c2e0-901">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-901">Requirements</span></span>

|<span data-ttu-id="4c2e0-902">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-902">Requirement</span></span>| <span data-ttu-id="4c2e0-903">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-904">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-905">1.0</span><span class="sxs-lookup"><span data-stu-id="4c2e0-905">1.0</span></span>|
|[<span data-ttu-id="4c2e0-906">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-907">ReadItem</span></span>|
|[<span data-ttu-id="4c2e0-908">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-909">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="4c2e0-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-910">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-910">Example</span></span>

<span data-ttu-id="4c2e0-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4c2e0-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4c2e0-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4c2e0-915">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4c2e0-p165">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-920">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-920">Parameters</span></span>

|<span data-ttu-id="4c2e0-921">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-921">Name</span></span>| <span data-ttu-id="4c2e0-922">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-922">Type</span></span>| <span data-ttu-id="4c2e0-923">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-923">Attributes</span></span>| <span data-ttu-id="4c2e0-924">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4c2e0-925">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-925">String</span></span>||<span data-ttu-id="4c2e0-926">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4c2e0-927">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="4c2e0-927">Object</span></span>| <span data-ttu-id="4c2e0-928">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-928">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-929">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4c2e0-930">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-930">Object</span></span>| <span data-ttu-id="4c2e0-931">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-931">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-932">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4c2e0-933">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-933">function</span></span>| <span data-ttu-id="4c2e0-934">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-934">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-935">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4c2e0-936">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4c2e0-937">エラー</span><span class="sxs-lookup"><span data-stu-id="4c2e0-937">Errors</span></span>

| <span data-ttu-id="4c2e0-938">エラー コード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-938">Error code</span></span> | <span data-ttu-id="4c2e0-939">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4c2e0-940">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-941">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-941">Requirements</span></span>

|<span data-ttu-id="4c2e0-942">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-942">Requirement</span></span>| <span data-ttu-id="4c2e0-943">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-944">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-945">1.1</span><span class="sxs-lookup"><span data-stu-id="4c2e0-945">1.1</span></span>|
|[<span data-ttu-id="4c2e0-946">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-946">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="4c2e0-948">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-948">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-949">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-950">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-950">Example</span></span>

<span data-ttu-id="4c2e0-951">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4c2e0-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4c2e0-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4c2e0-953">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4c2e0-p166">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c2e0-957">パラメーター</span><span class="sxs-lookup"><span data-stu-id="4c2e0-957">Parameters</span></span>

|<span data-ttu-id="4c2e0-958">名前</span><span class="sxs-lookup"><span data-stu-id="4c2e0-958">Name</span></span>| <span data-ttu-id="4c2e0-959">型</span><span class="sxs-lookup"><span data-stu-id="4c2e0-959">Type</span></span>| <span data-ttu-id="4c2e0-960">属性</span><span class="sxs-lookup"><span data-stu-id="4c2e0-960">Attributes</span></span>| <span data-ttu-id="4c2e0-961">説明</span><span class="sxs-lookup"><span data-stu-id="4c2e0-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4c2e0-962">String</span><span class="sxs-lookup"><span data-stu-id="4c2e0-962">String</span></span>||<span data-ttu-id="4c2e0-p167">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4c2e0-966">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-966">Object</span></span>| <span data-ttu-id="4c2e0-967">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-967">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-968">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4c2e0-969">Object</span><span class="sxs-lookup"><span data-stu-id="4c2e0-969">Object</span></span>| <span data-ttu-id="4c2e0-970">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-970">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-971">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="4c2e0-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4c2e0-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="4c2e0-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c2e0-973">&lt;optional&gt;</span></span>|<span data-ttu-id="4c2e0-p168">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4c2e0-p169">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4c2e0-978">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4c2e0-979">function</span><span class="sxs-lookup"><span data-stu-id="4c2e0-979">function</span></span>||<span data-ttu-id="4c2e0-980">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="4c2e0-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c2e0-981">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-981">Requirements</span></span>

|<span data-ttu-id="4c2e0-982">要件</span><span class="sxs-lookup"><span data-stu-id="4c2e0-982">Requirement</span></span>| <span data-ttu-id="4c2e0-983">値</span><span class="sxs-lookup"><span data-stu-id="4c2e0-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c2e0-984">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4c2e0-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c2e0-985">1.2</span><span class="sxs-lookup"><span data-stu-id="4c2e0-985">1.2</span></span>|
|[<span data-ttu-id="4c2e0-986">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4c2e0-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c2e0-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4c2e0-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="4c2e0-988">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4c2e0-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c2e0-989">作成</span><span class="sxs-lookup"><span data-stu-id="4c2e0-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4c2e0-990">例</span><span class="sxs-lookup"><span data-stu-id="4c2e0-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
