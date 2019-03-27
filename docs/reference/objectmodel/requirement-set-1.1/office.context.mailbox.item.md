---
title: Office. メールボックス-要件セット1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d3681f369570995c07256171fb6a65482648e85e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871620"
---
# <a name="item"></a><span data-ttu-id="aa117-102">item</span><span class="sxs-lookup"><span data-stu-id="aa117-102">item</span></span>

### <span data-ttu-id="aa117-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="aa117-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="aa117-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-107">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-107">Requirements</span></span>

|<span data-ttu-id="aa117-108">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-108">Requirement</span></span>| <span data-ttu-id="aa117-109">値</span><span class="sxs-lookup"><span data-stu-id="aa117-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-111">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-111">1.0</span></span>|
|[<span data-ttu-id="aa117-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="aa117-113">Restricted</span></span>|
|[<span data-ttu-id="aa117-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-115">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="aa117-116">例</span><span class="sxs-lookup"><span data-stu-id="aa117-116">Example</span></span>

<span data-ttu-id="aa117-117">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="aa117-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="aa117-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="aa117-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="aa117-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="aa117-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="aa117-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-122">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="aa117-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="aa117-123">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa117-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-124">型</span><span class="sxs-lookup"><span data-stu-id="aa117-124">Type</span></span>

*   <span data-ttu-id="aa117-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="aa117-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-126">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-126">Requirements</span></span>

|<span data-ttu-id="aa117-127">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-127">Requirement</span></span>| <span data-ttu-id="aa117-128">値</span><span class="sxs-lookup"><span data-stu-id="aa117-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-129">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-130">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-130">1.0</span></span>|
|[<span data-ttu-id="aa117-131">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-132">ReadItem</span></span>|
|[<span data-ttu-id="aa117-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-134">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-135">例</span><span class="sxs-lookup"><span data-stu-id="aa117-135">Example</span></span>

<span data-ttu-id="aa117-136">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="aa117-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="aa117-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="aa117-138">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="aa117-139">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-140">型</span><span class="sxs-lookup"><span data-stu-id="aa117-140">Type</span></span>

*   [<span data-ttu-id="aa117-141">受信者</span><span class="sxs-lookup"><span data-stu-id="aa117-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="aa117-142">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-142">Requirements</span></span>

|<span data-ttu-id="aa117-143">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-143">Requirement</span></span>| <span data-ttu-id="aa117-144">値</span><span class="sxs-lookup"><span data-stu-id="aa117-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-146">1.1</span><span class="sxs-lookup"><span data-stu-id="aa117-146">1.1</span></span>|
|[<span data-ttu-id="aa117-147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-148">ReadItem</span></span>|
|[<span data-ttu-id="aa117-149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-150">作成</span><span class="sxs-lookup"><span data-stu-id="aa117-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-151">例</span><span class="sxs-lookup"><span data-stu-id="aa117-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="aa117-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="aa117-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="aa117-153">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-154">型</span><span class="sxs-lookup"><span data-stu-id="aa117-154">Type</span></span>

*   [<span data-ttu-id="aa117-155">Body</span><span class="sxs-lookup"><span data-stu-id="aa117-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="aa117-156">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-156">Requirements</span></span>

|<span data-ttu-id="aa117-157">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-157">Requirement</span></span>| <span data-ttu-id="aa117-158">値</span><span class="sxs-lookup"><span data-stu-id="aa117-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-160">1.1</span><span class="sxs-lookup"><span data-stu-id="aa117-160">1.1</span></span>|
|[<span data-ttu-id="aa117-161">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-162">ReadItem</span></span>|
|[<span data-ttu-id="aa117-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-164">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-165">例</span><span class="sxs-lookup"><span data-stu-id="aa117-165">Example</span></span>

<span data-ttu-id="aa117-166">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="aa117-167">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="aa117-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="aa117-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="aa117-169">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa117-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="aa117-170">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="aa117-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-171">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-171">Read mode</span></span>

<span data-ttu-id="aa117-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-174">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-174">Compose mode</span></span>

<span data-ttu-id="aa117-175">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa117-176">型</span><span class="sxs-lookup"><span data-stu-id="aa117-176">Type</span></span>

*   <span data-ttu-id="aa117-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-178">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-178">Requirements</span></span>

|<span data-ttu-id="aa117-179">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-179">Requirement</span></span>| <span data-ttu-id="aa117-180">値</span><span class="sxs-lookup"><span data-stu-id="aa117-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-181">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-182">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-182">1.0</span></span>|
|[<span data-ttu-id="aa117-183">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-184">ReadItem</span></span>|
|[<span data-ttu-id="aa117-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="aa117-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="aa117-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="aa117-188">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="aa117-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="aa117-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="aa117-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-193">Type</span><span class="sxs-lookup"><span data-stu-id="aa117-193">Type</span></span>

*   <span data-ttu-id="aa117-194">String</span><span class="sxs-lookup"><span data-stu-id="aa117-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-195">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-195">Requirements</span></span>

|<span data-ttu-id="aa117-196">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-196">Requirement</span></span>| <span data-ttu-id="aa117-197">値</span><span class="sxs-lookup"><span data-stu-id="aa117-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-199">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-199">1.0</span></span>|
|[<span data-ttu-id="aa117-200">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-201">ReadItem</span></span>|
|[<span data-ttu-id="aa117-202">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-203">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-204">例</span><span class="sxs-lookup"><span data-stu-id="aa117-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="aa117-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="aa117-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="aa117-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-208">型</span><span class="sxs-lookup"><span data-stu-id="aa117-208">Type</span></span>

*   <span data-ttu-id="aa117-209">日付</span><span class="sxs-lookup"><span data-stu-id="aa117-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-210">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-210">Requirements</span></span>

|<span data-ttu-id="aa117-211">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-211">Requirement</span></span>| <span data-ttu-id="aa117-212">値</span><span class="sxs-lookup"><span data-stu-id="aa117-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-214">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-214">1.0</span></span>|
|[<span data-ttu-id="aa117-215">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-216">ReadItem</span></span>|
|[<span data-ttu-id="aa117-217">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-218">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-219">例</span><span class="sxs-lookup"><span data-stu-id="aa117-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="aa117-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="aa117-220">dateTimeModified :Date</span></span>

<span data-ttu-id="aa117-p111">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-223">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-224">型</span><span class="sxs-lookup"><span data-stu-id="aa117-224">Type</span></span>

*   <span data-ttu-id="aa117-225">日付</span><span class="sxs-lookup"><span data-stu-id="aa117-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-226">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-226">Requirements</span></span>

|<span data-ttu-id="aa117-227">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-227">Requirement</span></span>| <span data-ttu-id="aa117-228">値</span><span class="sxs-lookup"><span data-stu-id="aa117-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-229">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-230">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-230">1.0</span></span>|
|[<span data-ttu-id="aa117-231">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-232">ReadItem</span></span>|
|[<span data-ttu-id="aa117-233">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-234">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-235">例</span><span class="sxs-lookup"><span data-stu-id="aa117-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="aa117-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa117-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="aa117-237">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="aa117-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="aa117-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-240">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-240">Read mode</span></span>

<span data-ttu-id="aa117-241">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-242">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-242">Compose mode</span></span>

<span data-ttu-id="aa117-243">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="aa117-244">[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa117-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="aa117-245">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="aa117-246">型</span><span class="sxs-lookup"><span data-stu-id="aa117-246">Type</span></span>

*   <span data-ttu-id="aa117-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa117-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-248">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-248">Requirements</span></span>

|<span data-ttu-id="aa117-249">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-249">Requirement</span></span>| <span data-ttu-id="aa117-250">値</span><span class="sxs-lookup"><span data-stu-id="aa117-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-251">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-252">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-252">1.0</span></span>|
|[<span data-ttu-id="aa117-253">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-254">ReadItem</span></span>|
|[<span data-ttu-id="aa117-255">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-256">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="aa117-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="aa117-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="aa117-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="aa117-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-262">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="aa117-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-263">型</span><span class="sxs-lookup"><span data-stu-id="aa117-263">Type</span></span>

*   [<span data-ttu-id="aa117-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="aa117-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="aa117-265">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-265">Requirements</span></span>

|<span data-ttu-id="aa117-266">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-266">Requirement</span></span>| <span data-ttu-id="aa117-267">値</span><span class="sxs-lookup"><span data-stu-id="aa117-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-269">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-269">1.0</span></span>|
|[<span data-ttu-id="aa117-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-271">ReadItem</span></span>|
|[<span data-ttu-id="aa117-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-273">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-274">例</span><span class="sxs-lookup"><span data-stu-id="aa117-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="aa117-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="aa117-275">internetMessageId :String</span></span>

<span data-ttu-id="aa117-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-278">Type</span><span class="sxs-lookup"><span data-stu-id="aa117-278">Type</span></span>

*   <span data-ttu-id="aa117-279">String</span><span class="sxs-lookup"><span data-stu-id="aa117-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-280">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-280">Requirements</span></span>

|<span data-ttu-id="aa117-281">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-281">Requirement</span></span>| <span data-ttu-id="aa117-282">値</span><span class="sxs-lookup"><span data-stu-id="aa117-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-283">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-284">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-284">1.0</span></span>|
|[<span data-ttu-id="aa117-285">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-286">ReadItem</span></span>|
|[<span data-ttu-id="aa117-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-288">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-289">例</span><span class="sxs-lookup"><span data-stu-id="aa117-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="aa117-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="aa117-290">itemClass :String</span></span>

<span data-ttu-id="aa117-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="aa117-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="aa117-295">型</span><span class="sxs-lookup"><span data-stu-id="aa117-295">Type</span></span> | <span data-ttu-id="aa117-296">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-296">Description</span></span> | <span data-ttu-id="aa117-297">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="aa117-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="aa117-298">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="aa117-298">Appointment items</span></span> | <span data-ttu-id="aa117-299">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="aa117-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="aa117-300">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="aa117-300">Message items</span></span> | <span data-ttu-id="aa117-301">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="aa117-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="aa117-302">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-303">型</span><span class="sxs-lookup"><span data-stu-id="aa117-303">Type</span></span>

*   <span data-ttu-id="aa117-304">String</span><span class="sxs-lookup"><span data-stu-id="aa117-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-305">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-305">Requirements</span></span>

|<span data-ttu-id="aa117-306">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-306">Requirement</span></span>| <span data-ttu-id="aa117-307">値</span><span class="sxs-lookup"><span data-stu-id="aa117-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-308">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-309">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-309">1.0</span></span>|
|[<span data-ttu-id="aa117-310">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-311">ReadItem</span></span>|
|[<span data-ttu-id="aa117-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-313">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-314">例</span><span class="sxs-lookup"><span data-stu-id="aa117-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="aa117-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="aa117-315">(nullable) itemId :String</span></span>

<span data-ttu-id="aa117-p118">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-318">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="aa117-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="aa117-319">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="aa117-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="aa117-320">この値を使用して REST API を呼び出す前に、を`Office.context.mailbox.convertToRestId`使用して変換する必要があります。これは、要件セット1.3 から開始できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="aa117-321">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aa117-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-322">Type</span><span class="sxs-lookup"><span data-stu-id="aa117-322">Type</span></span>

*   <span data-ttu-id="aa117-323">String</span><span class="sxs-lookup"><span data-stu-id="aa117-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-324">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-324">Requirements</span></span>

|<span data-ttu-id="aa117-325">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-325">Requirement</span></span>| <span data-ttu-id="aa117-326">値</span><span class="sxs-lookup"><span data-stu-id="aa117-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-327">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-328">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-328">1.0</span></span>|
|[<span data-ttu-id="aa117-329">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-330">ReadItem</span></span>|
|[<span data-ttu-id="aa117-331">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-332">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-333">例</span><span class="sxs-lookup"><span data-stu-id="aa117-333">Example</span></span>

<span data-ttu-id="aa117-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="aa117-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="aa117-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="aa117-337">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="aa117-338">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="aa117-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-339">型</span><span class="sxs-lookup"><span data-stu-id="aa117-339">Type</span></span>

*   [<span data-ttu-id="aa117-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="aa117-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="aa117-341">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-341">Requirements</span></span>

|<span data-ttu-id="aa117-342">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-342">Requirement</span></span>| <span data-ttu-id="aa117-343">値</span><span class="sxs-lookup"><span data-stu-id="aa117-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-344">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-345">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-345">1.0</span></span>|
|[<span data-ttu-id="aa117-346">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-347">ReadItem</span></span>|
|[<span data-ttu-id="aa117-348">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-349">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-350">例</span><span class="sxs-lookup"><span data-stu-id="aa117-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="aa117-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="aa117-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="aa117-352">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-353">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-353">Read mode</span></span>

<span data-ttu-id="aa117-354">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-355">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-355">Compose mode</span></span>

<span data-ttu-id="aa117-356">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa117-357">型</span><span class="sxs-lookup"><span data-stu-id="aa117-357">Type</span></span>

*   <span data-ttu-id="aa117-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="aa117-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-359">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-359">Requirements</span></span>

|<span data-ttu-id="aa117-360">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-360">Requirement</span></span>| <span data-ttu-id="aa117-361">値</span><span class="sxs-lookup"><span data-stu-id="aa117-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-362">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-363">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-363">1.0</span></span>|
|[<span data-ttu-id="aa117-364">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-365">ReadItem</span></span>|
|[<span data-ttu-id="aa117-366">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-367">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="aa117-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="aa117-368">normalizedSubject :String</span></span>

<span data-ttu-id="aa117-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="aa117-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-373">Type</span><span class="sxs-lookup"><span data-stu-id="aa117-373">Type</span></span>

*   <span data-ttu-id="aa117-374">String</span><span class="sxs-lookup"><span data-stu-id="aa117-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-375">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-375">Requirements</span></span>

|<span data-ttu-id="aa117-376">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-376">Requirement</span></span>| <span data-ttu-id="aa117-377">値</span><span class="sxs-lookup"><span data-stu-id="aa117-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-378">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-379">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-379">1.0</span></span>|
|[<span data-ttu-id="aa117-380">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-381">ReadItem</span></span>|
|[<span data-ttu-id="aa117-382">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-383">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-384">例</span><span class="sxs-lookup"><span data-stu-id="aa117-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="aa117-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="aa117-386">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa117-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="aa117-387">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="aa117-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-388">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-388">Read mode</span></span>

<span data-ttu-id="aa117-389">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-390">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-390">Compose mode</span></span>

<span data-ttu-id="aa117-391">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa117-392">型</span><span class="sxs-lookup"><span data-stu-id="aa117-392">Type</span></span>

*   <span data-ttu-id="aa117-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-394">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-394">Requirements</span></span>

|<span data-ttu-id="aa117-395">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-395">Requirement</span></span>| <span data-ttu-id="aa117-396">値</span><span class="sxs-lookup"><span data-stu-id="aa117-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-397">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-398">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-398">1.0</span></span>|
|[<span data-ttu-id="aa117-399">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-400">ReadItem</span></span>|
|[<span data-ttu-id="aa117-401">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-402">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="aa117-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="aa117-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="aa117-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-406">型</span><span class="sxs-lookup"><span data-stu-id="aa117-406">Type</span></span>

*   [<span data-ttu-id="aa117-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="aa117-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="aa117-408">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-408">Requirements</span></span>

|<span data-ttu-id="aa117-409">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-409">Requirement</span></span>| <span data-ttu-id="aa117-410">値</span><span class="sxs-lookup"><span data-stu-id="aa117-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-411">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-412">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-412">1.0</span></span>|
|[<span data-ttu-id="aa117-413">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-414">ReadItem</span></span>|
|[<span data-ttu-id="aa117-415">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-416">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-417">例</span><span class="sxs-lookup"><span data-stu-id="aa117-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="aa117-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="aa117-419">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa117-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="aa117-420">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="aa117-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-421">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-421">Read mode</span></span>

<span data-ttu-id="aa117-422">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-423">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-423">Compose mode</span></span>

<span data-ttu-id="aa117-424">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="aa117-425">型</span><span class="sxs-lookup"><span data-stu-id="aa117-425">Type</span></span>

*   <span data-ttu-id="aa117-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-427">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-427">Requirements</span></span>

|<span data-ttu-id="aa117-428">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-428">Requirement</span></span>| <span data-ttu-id="aa117-429">値</span><span class="sxs-lookup"><span data-stu-id="aa117-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-430">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-431">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-431">1.0</span></span>|
|[<span data-ttu-id="aa117-432">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-433">ReadItem</span></span>|
|[<span data-ttu-id="aa117-434">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-435">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="aa117-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="aa117-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="aa117-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="aa117-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="aa117-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-441">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="aa117-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="aa117-442">型</span><span class="sxs-lookup"><span data-stu-id="aa117-442">Type</span></span>

*   [<span data-ttu-id="aa117-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="aa117-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="aa117-444">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-444">Requirements</span></span>

|<span data-ttu-id="aa117-445">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-445">Requirement</span></span>| <span data-ttu-id="aa117-446">値</span><span class="sxs-lookup"><span data-stu-id="aa117-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-448">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-448">1.0</span></span>|
|[<span data-ttu-id="aa117-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-450">ReadItem</span></span>|
|[<span data-ttu-id="aa117-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-453">例</span><span class="sxs-lookup"><span data-stu-id="aa117-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="aa117-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa117-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="aa117-455">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="aa117-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="aa117-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-458">Read mode</span></span>

<span data-ttu-id="aa117-459">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-460">Compose mode</span></span>

<span data-ttu-id="aa117-461">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="aa117-462">[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="aa117-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="aa117-463">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="aa117-464">型</span><span class="sxs-lookup"><span data-stu-id="aa117-464">Type</span></span>

*   <span data-ttu-id="aa117-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="aa117-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-466">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-466">Requirements</span></span>

|<span data-ttu-id="aa117-467">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-467">Requirement</span></span>| <span data-ttu-id="aa117-468">値</span><span class="sxs-lookup"><span data-stu-id="aa117-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-469">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-470">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-470">1.0</span></span>|
|[<span data-ttu-id="aa117-471">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-472">ReadItem</span></span>|
|[<span data-ttu-id="aa117-473">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-474">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="aa117-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="aa117-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="aa117-476">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="aa117-477">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="aa117-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-478">Read mode</span></span>

<span data-ttu-id="aa117-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-481">Compose mode</span></span>

<span data-ttu-id="aa117-482">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="aa117-483">型</span><span class="sxs-lookup"><span data-stu-id="aa117-483">Type</span></span>

*   <span data-ttu-id="aa117-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="aa117-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-485">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-485">Requirements</span></span>

|<span data-ttu-id="aa117-486">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-486">Requirement</span></span>| <span data-ttu-id="aa117-487">値</span><span class="sxs-lookup"><span data-stu-id="aa117-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-489">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-489">1.0</span></span>|
|[<span data-ttu-id="aa117-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-491">ReadItem</span></span>|
|[<span data-ttu-id="aa117-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-493">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="aa117-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="aa117-495">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="aa117-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="aa117-496">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="aa117-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="aa117-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="aa117-497">Read mode</span></span>

<span data-ttu-id="aa117-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="aa117-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="aa117-500">Compose mode</span></span>

<span data-ttu-id="aa117-501">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="aa117-502">型</span><span class="sxs-lookup"><span data-stu-id="aa117-502">Type</span></span>

*   <span data-ttu-id="aa117-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="aa117-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-504">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-504">Requirements</span></span>

|<span data-ttu-id="aa117-505">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-505">Requirement</span></span>| <span data-ttu-id="aa117-506">値</span><span class="sxs-lookup"><span data-stu-id="aa117-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-508">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-508">1.0</span></span>|
|[<span data-ttu-id="aa117-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-510">ReadItem</span></span>|
|[<span data-ttu-id="aa117-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-512">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="aa117-513">メソッド</span><span class="sxs-lookup"><span data-stu-id="aa117-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="aa117-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa117-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="aa117-515">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="aa117-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="aa117-516">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="aa117-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="aa117-517">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-518">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-518">Parameters</span></span>

|<span data-ttu-id="aa117-519">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-519">Name</span></span>| <span data-ttu-id="aa117-520">種類</span><span class="sxs-lookup"><span data-stu-id="aa117-520">Type</span></span>| <span data-ttu-id="aa117-521">属性</span><span class="sxs-lookup"><span data-stu-id="aa117-521">Attributes</span></span>| <span data-ttu-id="aa117-522">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="aa117-523">String</span><span class="sxs-lookup"><span data-stu-id="aa117-523">String</span></span>||<span data-ttu-id="aa117-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="aa117-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="aa117-526">String</span><span class="sxs-lookup"><span data-stu-id="aa117-526">String</span></span>||<span data-ttu-id="aa117-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="aa117-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="aa117-529">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-529">Object</span></span>| <span data-ttu-id="aa117-530">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-530">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-531">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="aa117-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="aa117-532">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-532">Object</span></span>| <span data-ttu-id="aa117-533">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-533">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-534">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="aa117-535">function</span><span class="sxs-lookup"><span data-stu-id="aa117-535">function</span></span>| <span data-ttu-id="aa117-536">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-536">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-537">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa117-538">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="aa117-539">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="aa117-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa117-540">エラー</span><span class="sxs-lookup"><span data-stu-id="aa117-540">Errors</span></span>

| <span data-ttu-id="aa117-541">エラー コード</span><span class="sxs-lookup"><span data-stu-id="aa117-541">Error code</span></span> | <span data-ttu-id="aa117-542">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="aa117-543">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="aa117-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="aa117-544">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="aa117-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="aa117-545">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="aa117-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa117-546">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-546">Requirements</span></span>

|<span data-ttu-id="aa117-547">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-547">Requirement</span></span>| <span data-ttu-id="aa117-548">値</span><span class="sxs-lookup"><span data-stu-id="aa117-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-549">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-550">1.1</span><span class="sxs-lookup"><span data-stu-id="aa117-550">1.1</span></span>|
|[<span data-ttu-id="aa117-551">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa117-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa117-553">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-554">作成</span><span class="sxs-lookup"><span data-stu-id="aa117-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-555">例</span><span class="sxs-lookup"><span data-stu-id="aa117-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="aa117-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa117-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="aa117-557">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="aa117-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="aa117-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="aa117-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="aa117-561">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="aa117-562">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="aa117-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-563">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-563">Parameters</span></span>

|<span data-ttu-id="aa117-564">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-564">Name</span></span>| <span data-ttu-id="aa117-565">型</span><span class="sxs-lookup"><span data-stu-id="aa117-565">Type</span></span>| <span data-ttu-id="aa117-566">属性</span><span class="sxs-lookup"><span data-stu-id="aa117-566">Attributes</span></span>| <span data-ttu-id="aa117-567">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="aa117-568">String</span><span class="sxs-lookup"><span data-stu-id="aa117-568">String</span></span>||<span data-ttu-id="aa117-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="aa117-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="aa117-571">String</span><span class="sxs-lookup"><span data-stu-id="aa117-571">String</span></span>||<span data-ttu-id="aa117-572">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="aa117-572">The subject of the item to be attached.</span></span> <span data-ttu-id="aa117-573">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="aa117-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="aa117-574">Object</span><span class="sxs-lookup"><span data-stu-id="aa117-574">Object</span></span>| <span data-ttu-id="aa117-575">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-575">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-576">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="aa117-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="aa117-577">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-577">Object</span></span>| <span data-ttu-id="aa117-578">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-578">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-579">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="aa117-580">function</span><span class="sxs-lookup"><span data-stu-id="aa117-580">function</span></span>| <span data-ttu-id="aa117-581">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-581">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-582">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa117-583">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="aa117-584">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="aa117-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa117-585">エラー</span><span class="sxs-lookup"><span data-stu-id="aa117-585">Errors</span></span>

| <span data-ttu-id="aa117-586">エラー コード</span><span class="sxs-lookup"><span data-stu-id="aa117-586">Error code</span></span> | <span data-ttu-id="aa117-587">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="aa117-588">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="aa117-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa117-589">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-589">Requirements</span></span>

|<span data-ttu-id="aa117-590">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-590">Requirement</span></span>| <span data-ttu-id="aa117-591">値</span><span class="sxs-lookup"><span data-stu-id="aa117-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-592">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-593">1.1</span><span class="sxs-lookup"><span data-stu-id="aa117-593">1.1</span></span>|
|[<span data-ttu-id="aa117-594">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa117-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa117-596">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-597">作成</span><span class="sxs-lookup"><span data-stu-id="aa117-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-598">例</span><span class="sxs-lookup"><span data-stu-id="aa117-598">Example</span></span>

<span data-ttu-id="aa117-599">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="aa117-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="aa117-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="aa117-601">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-602">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa117-603">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="aa117-604">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="aa117-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-605">へ`displayReplyAllForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="aa117-606">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="aa117-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-607">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-607">Parameters</span></span>

|<span data-ttu-id="aa117-608">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-608">Name</span></span>| <span data-ttu-id="aa117-609">型</span><span class="sxs-lookup"><span data-stu-id="aa117-609">Type</span></span>| <span data-ttu-id="aa117-610">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="aa117-611">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="aa117-611">String &#124; Object</span></span>| |<span data-ttu-id="aa117-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="aa117-614">**または**</span><span class="sxs-lookup"><span data-stu-id="aa117-614">**OR**</span></span><br/><span data-ttu-id="aa117-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="aa117-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="aa117-617">String</span><span class="sxs-lookup"><span data-stu-id="aa117-617">String</span></span> | <span data-ttu-id="aa117-618">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-618">&lt;optional&gt;</span></span> | <span data-ttu-id="aa117-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="aa117-621">function</span><span class="sxs-lookup"><span data-stu-id="aa117-621">function</span></span> | <span data-ttu-id="aa117-622">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-622">&lt;optional&gt;</span></span> | <span data-ttu-id="aa117-623">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa117-624">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-624">Requirements</span></span>

|<span data-ttu-id="aa117-625">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-625">Requirement</span></span>| <span data-ttu-id="aa117-626">値</span><span class="sxs-lookup"><span data-stu-id="aa117-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-627">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-628">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-628">1.0</span></span>|
|[<span data-ttu-id="aa117-629">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-630">ReadItem</span></span>|
|[<span data-ttu-id="aa117-631">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-632">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa117-633">例</span><span class="sxs-lookup"><span data-stu-id="aa117-633">Examples</span></span>

<span data-ttu-id="aa117-634">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="aa117-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="aa117-635">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="aa117-636">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="aa117-637">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-637">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="aa117-638">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="aa117-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="aa117-639">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-640">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-640">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa117-641">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-641">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="aa117-642">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="aa117-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-643">へ`displayReplyForm`の呼び出しに添付ファイルを含める機能は、要件セット1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="aa117-644">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="aa117-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-645">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-645">Parameters</span></span>

|<span data-ttu-id="aa117-646">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-646">Name</span></span>| <span data-ttu-id="aa117-647">型</span><span class="sxs-lookup"><span data-stu-id="aa117-647">Type</span></span>| <span data-ttu-id="aa117-648">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="aa117-649">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="aa117-649">String &#124; Object</span></span>| | <span data-ttu-id="aa117-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="aa117-652">**または**</span><span class="sxs-lookup"><span data-stu-id="aa117-652">**OR**</span></span><br/><span data-ttu-id="aa117-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="aa117-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="aa117-655">String</span><span class="sxs-lookup"><span data-stu-id="aa117-655">String</span></span> | <span data-ttu-id="aa117-656">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-656">&lt;optional&gt;</span></span> | <span data-ttu-id="aa117-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="aa117-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="aa117-659">function</span><span class="sxs-lookup"><span data-stu-id="aa117-659">function</span></span> | <span data-ttu-id="aa117-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-660">&lt;optional&gt;</span></span> | <span data-ttu-id="aa117-661">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa117-662">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-662">Requirements</span></span>

|<span data-ttu-id="aa117-663">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-663">Requirement</span></span>| <span data-ttu-id="aa117-664">値</span><span class="sxs-lookup"><span data-stu-id="aa117-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-665">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-666">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-666">1.0</span></span>|
|[<span data-ttu-id="aa117-667">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-668">ReadItem</span></span>|
|[<span data-ttu-id="aa117-669">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-670">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="aa117-671">例</span><span class="sxs-lookup"><span data-stu-id="aa117-671">Examples</span></span>

<span data-ttu-id="aa117-672">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="aa117-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="aa117-673">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="aa117-674">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="aa117-675">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="aa117-675">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="aa117-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="aa117-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="aa117-677">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-678">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-678">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-679">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-679">Requirements</span></span>

|<span data-ttu-id="aa117-680">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-680">Requirement</span></span>| <span data-ttu-id="aa117-681">値</span><span class="sxs-lookup"><span data-stu-id="aa117-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-682">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-683">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-683">1.0</span></span>|
|[<span data-ttu-id="aa117-684">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-684">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-685">ReadItem</span></span>|
|[<span data-ttu-id="aa117-686">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-686">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-687">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa117-688">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa117-688">Returns:</span></span>

<span data-ttu-id="aa117-689">型:[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="aa117-689">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="aa117-690">例</span><span class="sxs-lookup"><span data-stu-id="aa117-690">Example</span></span>

<span data-ttu-id="aa117-691">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="aa117-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="aa117-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="aa117-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="aa117-693">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="aa117-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-694">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-694">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-695">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-695">Parameters</span></span>

|<span data-ttu-id="aa117-696">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-696">Name</span></span>| <span data-ttu-id="aa117-697">型</span><span class="sxs-lookup"><span data-stu-id="aa117-697">Type</span></span>| <span data-ttu-id="aa117-698">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="aa117-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="aa117-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="aa117-700">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="aa117-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa117-701">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa117-701">Requirements</span></span>

|<span data-ttu-id="aa117-702">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-702">Requirement</span></span>| <span data-ttu-id="aa117-703">値</span><span class="sxs-lookup"><span data-stu-id="aa117-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-704">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-705">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-705">1.0</span></span>|
|[<span data-ttu-id="aa117-706">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-706">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-707">制限あり</span><span class="sxs-lookup"><span data-stu-id="aa117-707">Restricted</span></span>|
|[<span data-ttu-id="aa117-708">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-708">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-709">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa117-710">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa117-710">Returns:</span></span>

<span data-ttu-id="aa117-711">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="aa117-712">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="aa117-713">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="aa117-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="aa117-714">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="aa117-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="aa117-715">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="aa117-715">Value of `entityType`</span></span> | <span data-ttu-id="aa117-716">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="aa117-716">Type of objects in returned array</span></span> | <span data-ttu-id="aa117-717">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="aa117-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="aa117-718">文字列</span><span class="sxs-lookup"><span data-stu-id="aa117-718">String</span></span> | <span data-ttu-id="aa117-719">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="aa117-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="aa117-720">連絡先</span><span class="sxs-lookup"><span data-stu-id="aa117-720">Contact</span></span> | <span data-ttu-id="aa117-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa117-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="aa117-722">文字列</span><span class="sxs-lookup"><span data-stu-id="aa117-722">String</span></span> | <span data-ttu-id="aa117-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa117-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="aa117-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="aa117-724">MeetingSuggestion</span></span> | <span data-ttu-id="aa117-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa117-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="aa117-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="aa117-726">PhoneNumber</span></span> | <span data-ttu-id="aa117-727">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="aa117-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="aa117-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="aa117-728">TaskSuggestion</span></span> | <span data-ttu-id="aa117-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="aa117-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="aa117-730">文字列</span><span class="sxs-lookup"><span data-stu-id="aa117-730">String</span></span> | <span data-ttu-id="aa117-731">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="aa117-731">**Restricted**</span></span> |

<span data-ttu-id="aa117-732">型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="aa117-732">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="aa117-733">例</span><span class="sxs-lookup"><span data-stu-id="aa117-733">Example</span></span>

<span data-ttu-id="aa117-734">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="aa117-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="aa117-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="aa117-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="aa117-736">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-737">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-737">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa117-738">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-739">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-739">Parameters</span></span>

|<span data-ttu-id="aa117-740">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-740">Name</span></span>| <span data-ttu-id="aa117-741">型</span><span class="sxs-lookup"><span data-stu-id="aa117-741">Type</span></span>| <span data-ttu-id="aa117-742">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="aa117-743">String</span><span class="sxs-lookup"><span data-stu-id="aa117-743">String</span></span>|<span data-ttu-id="aa117-744">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="aa117-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa117-745">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-745">Requirements</span></span>

|<span data-ttu-id="aa117-746">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-746">Requirement</span></span>| <span data-ttu-id="aa117-747">値</span><span class="sxs-lookup"><span data-stu-id="aa117-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-749">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-749">1.0</span></span>|
|[<span data-ttu-id="aa117-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-751">ReadItem</span></span>|
|[<span data-ttu-id="aa117-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-753">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa117-754">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa117-754">Returns:</span></span>

<span data-ttu-id="aa117-p146">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="aa117-757">型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="aa117-757">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="aa117-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="aa117-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="aa117-759">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-760">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa117-p147">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="aa117-764">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="aa117-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="aa117-765">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="aa117-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="aa117-p148">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="aa117-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa117-768">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-768">Requirements</span></span>

|<span data-ttu-id="aa117-769">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-769">Requirement</span></span>| <span data-ttu-id="aa117-770">値</span><span class="sxs-lookup"><span data-stu-id="aa117-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-771">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-772">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-772">1.0</span></span>|
|[<span data-ttu-id="aa117-773">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-773">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-774">ReadItem</span></span>|
|[<span data-ttu-id="aa117-775">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-775">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-776">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa117-777">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa117-777">Returns:</span></span>

<span data-ttu-id="aa117-p149">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="aa117-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="aa117-780">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa117-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa117-781">Object</span><span class="sxs-lookup"><span data-stu-id="aa117-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="aa117-782">例</span><span class="sxs-lookup"><span data-stu-id="aa117-782">Example</span></span>

<span data-ttu-id="aa117-783">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="aa117-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="aa117-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="aa117-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="aa117-785">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="aa117-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="aa117-786">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="aa117-786">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="aa117-787">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="aa117-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="aa117-p150">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="aa117-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-790">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-790">Parameters</span></span>

|<span data-ttu-id="aa117-791">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-791">Name</span></span>| <span data-ttu-id="aa117-792">型</span><span class="sxs-lookup"><span data-stu-id="aa117-792">Type</span></span>| <span data-ttu-id="aa117-793">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="aa117-794">String</span><span class="sxs-lookup"><span data-stu-id="aa117-794">String</span></span>|<span data-ttu-id="aa117-795">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="aa117-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa117-796">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-796">Requirements</span></span>

|<span data-ttu-id="aa117-797">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-797">Requirement</span></span>| <span data-ttu-id="aa117-798">値</span><span class="sxs-lookup"><span data-stu-id="aa117-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-799">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-800">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-800">1.0</span></span>|
|[<span data-ttu-id="aa117-801">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-802">ReadItem</span></span>|
|[<span data-ttu-id="aa117-803">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-804">読み取り</span><span class="sxs-lookup"><span data-stu-id="aa117-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa117-805">戻り値:</span><span class="sxs-lookup"><span data-stu-id="aa117-805">Returns:</span></span>

<span data-ttu-id="aa117-806">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="aa117-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="aa117-807">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="aa117-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="aa117-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="aa117-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="aa117-809">例</span><span class="sxs-lookup"><span data-stu-id="aa117-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="aa117-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa117-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa117-811">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="aa117-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="aa117-p151">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="aa117-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-815">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-815">Parameters</span></span>

|<span data-ttu-id="aa117-816">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-816">Name</span></span>| <span data-ttu-id="aa117-817">型</span><span class="sxs-lookup"><span data-stu-id="aa117-817">Type</span></span>| <span data-ttu-id="aa117-818">属性</span><span class="sxs-lookup"><span data-stu-id="aa117-818">Attributes</span></span>| <span data-ttu-id="aa117-819">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="aa117-820">function</span><span class="sxs-lookup"><span data-stu-id="aa117-820">function</span></span>||<span data-ttu-id="aa117-821">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa117-822">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="aa117-823">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="aa117-824">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-824">Object</span></span>| <span data-ttu-id="aa117-825">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-825">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-826">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="aa117-827">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="aa117-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa117-828">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-828">Requirements</span></span>

|<span data-ttu-id="aa117-829">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-829">Requirement</span></span>| <span data-ttu-id="aa117-830">値</span><span class="sxs-lookup"><span data-stu-id="aa117-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-831">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-832">1.0</span><span class="sxs-lookup"><span data-stu-id="aa117-832">1.0</span></span>|
|[<span data-ttu-id="aa117-833">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-833">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa117-834">ReadItem</span></span>|
|[<span data-ttu-id="aa117-835">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-835">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-836">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="aa117-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-837">例</span><span class="sxs-lookup"><span data-stu-id="aa117-837">Example</span></span>

<span data-ttu-id="aa117-p154">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="aa117-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa117-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="aa117-842">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="aa117-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="aa117-p155">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="aa117-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa117-847">パラメーター</span><span class="sxs-lookup"><span data-stu-id="aa117-847">Parameters</span></span>

|<span data-ttu-id="aa117-848">名前</span><span class="sxs-lookup"><span data-stu-id="aa117-848">Name</span></span>| <span data-ttu-id="aa117-849">型</span><span class="sxs-lookup"><span data-stu-id="aa117-849">Type</span></span>| <span data-ttu-id="aa117-850">属性</span><span class="sxs-lookup"><span data-stu-id="aa117-850">Attributes</span></span>| <span data-ttu-id="aa117-851">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="aa117-852">String</span><span class="sxs-lookup"><span data-stu-id="aa117-852">String</span></span>||<span data-ttu-id="aa117-853">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="aa117-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="aa117-854">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-854">Object</span></span>| <span data-ttu-id="aa117-855">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-855">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-856">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="aa117-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="aa117-857">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="aa117-857">Object</span></span>| <span data-ttu-id="aa117-858">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-858">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-859">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="aa117-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="aa117-860">function</span><span class="sxs-lookup"><span data-stu-id="aa117-860">function</span></span>| <span data-ttu-id="aa117-861">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="aa117-861">&lt;optional&gt;</span></span>|<span data-ttu-id="aa117-862">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="aa117-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="aa117-863">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="aa117-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa117-864">エラー</span><span class="sxs-lookup"><span data-stu-id="aa117-864">Errors</span></span>

| <span data-ttu-id="aa117-865">エラー コード</span><span class="sxs-lookup"><span data-stu-id="aa117-865">Error code</span></span> | <span data-ttu-id="aa117-866">説明</span><span class="sxs-lookup"><span data-stu-id="aa117-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="aa117-867">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="aa117-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa117-868">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-868">Requirements</span></span>

|<span data-ttu-id="aa117-869">要件</span><span class="sxs-lookup"><span data-stu-id="aa117-869">Requirement</span></span>| <span data-ttu-id="aa117-870">値</span><span class="sxs-lookup"><span data-stu-id="aa117-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa117-871">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="aa117-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa117-872">1.1</span><span class="sxs-lookup"><span data-stu-id="aa117-872">1.1</span></span>|
|[<span data-ttu-id="aa117-873">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="aa117-873">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa117-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="aa117-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="aa117-875">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="aa117-875">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa117-876">作成</span><span class="sxs-lookup"><span data-stu-id="aa117-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="aa117-877">例</span><span class="sxs-lookup"><span data-stu-id="aa117-877">Example</span></span>

<span data-ttu-id="aa117-878">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="aa117-878">The following code removes an attachment with an identifier of '0'.</span></span>

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
