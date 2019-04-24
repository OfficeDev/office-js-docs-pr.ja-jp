---
title: Office. メールボックス-要件セット1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd8e71e39940fcf0de50982ef1cdb6825abb7221
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450367"
---
# <a name="item"></a><span data-ttu-id="dd0eb-102">item</span><span class="sxs-lookup"><span data-stu-id="dd0eb-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="dd0eb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="dd0eb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="dd0eb-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-106">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-106">Requirements</span></span>

|<span data-ttu-id="dd0eb-107">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-107">Requirement</span></span>| <span data-ttu-id="dd0eb-108">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-110">1.0</span></span>|
|[<span data-ttu-id="dd0eb-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="dd0eb-112">Restricted</span></span>|
|[<span data-ttu-id="dd0eb-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-114">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="dd0eb-115">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-115">Example</span></span>

<span data-ttu-id="dd0eb-116">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="dd0eb-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="dd0eb-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="dd0eb-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="dd0eb-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="dd0eb-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-121">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="dd0eb-122">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-123">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-123">Type</span></span>

*   <span data-ttu-id="dd0eb-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="dd0eb-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-125">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-125">Requirements</span></span>

|<span data-ttu-id="dd0eb-126">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-126">Requirement</span></span>| <span data-ttu-id="dd0eb-127">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-128">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-129">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-129">1.0</span></span>|
|[<span data-ttu-id="dd0eb-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-131">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-134">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-134">Example</span></span>

<span data-ttu-id="dd0eb-135">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="dd0eb-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="dd0eb-137">メッセージの Bcc (ブラインドカーボンコピー) 行を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="dd0eb-138">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-139">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-139">Type</span></span>

*   [<span data-ttu-id="dd0eb-140">受信者</span><span class="sxs-lookup"><span data-stu-id="dd0eb-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-141">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-141">Requirements</span></span>

|<span data-ttu-id="dd0eb-142">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-142">Requirement</span></span>| <span data-ttu-id="dd0eb-143">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-145">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0eb-145">1.1</span></span>|
|[<span data-ttu-id="dd0eb-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-147">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-149">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-150">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="dd0eb-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="dd0eb-152">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-153">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-153">Type</span></span>

*   [<span data-ttu-id="dd0eb-154">Body</span><span class="sxs-lookup"><span data-stu-id="dd0eb-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-155">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-155">Requirements</span></span>

|<span data-ttu-id="dd0eb-156">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-156">Requirement</span></span>| <span data-ttu-id="dd0eb-157">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-159">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0eb-159">1.1</span></span>|
|[<span data-ttu-id="dd0eb-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-161">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-164">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-164">Example</span></span>

<span data-ttu-id="dd0eb-165">この例では、メッセージの本文をプレーン テキストで取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="dd0eb-166">次の例は、コールバック関数に渡される結果パラメーターの例です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="dd0eb-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-167">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="dd0eb-168">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="dd0eb-169">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-170">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-170">Read mode</span></span>

<span data-ttu-id="dd0eb-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-173">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-173">Compose mode</span></span>

<span data-ttu-id="dd0eb-174">`cc` プロパティは、メッセージの **Cc** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-175">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-175">Type</span></span>

*   <span data-ttu-id="dd0eb-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-177">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-177">Requirements</span></span>

|<span data-ttu-id="dd0eb-178">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-178">Requirement</span></span>| <span data-ttu-id="dd0eb-179">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-181">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-181">1.0</span></span>|
|[<span data-ttu-id="dd0eb-182">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-183">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-184">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-185">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-185">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="dd0eb-186">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-186">(nullable) conversationId :String</span></span>

<span data-ttu-id="dd0eb-187">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="dd0eb-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="dd0eb-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-192">Type</span><span class="sxs-lookup"><span data-stu-id="dd0eb-192">Type</span></span>

*   <span data-ttu-id="dd0eb-193">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-194">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-194">Requirements</span></span>

|<span data-ttu-id="dd0eb-195">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-195">Requirement</span></span>| <span data-ttu-id="dd0eb-196">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-198">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-198">1.0</span></span>|
|[<span data-ttu-id="dd0eb-199">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-200">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-203">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="dd0eb-204">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="dd0eb-204">dateTimeCreated :Date</span></span>

<span data-ttu-id="dd0eb-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-207">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-207">Type</span></span>

*   <span data-ttu-id="dd0eb-208">日付</span><span class="sxs-lookup"><span data-stu-id="dd0eb-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-209">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-209">Requirements</span></span>

|<span data-ttu-id="dd0eb-210">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-210">Requirement</span></span>| <span data-ttu-id="dd0eb-211">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-213">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-213">1.0</span></span>|
|[<span data-ttu-id="dd0eb-214">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-215">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-216">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-217">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-218">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="dd0eb-219">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="dd0eb-219">dateTimeModified :Date</span></span>

<span data-ttu-id="dd0eb-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-222">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-223">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-223">Type</span></span>

*   <span data-ttu-id="dd0eb-224">日付</span><span class="sxs-lookup"><span data-stu-id="dd0eb-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-225">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-225">Requirements</span></span>

|<span data-ttu-id="dd0eb-226">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-226">Requirement</span></span>| <span data-ttu-id="dd0eb-227">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-228">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-229">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-229">1.0</span></span>|
|[<span data-ttu-id="dd0eb-230">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-231">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-232">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-233">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-234">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="dd0eb-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-235">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="dd0eb-236">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="dd0eb-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-239">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-239">Read mode</span></span>

<span data-ttu-id="dd0eb-240">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-241">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-241">Compose mode</span></span>

<span data-ttu-id="dd0eb-242">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="dd0eb-243">[`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="dd0eb-244">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="dd0eb-245">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-245">Type</span></span>

*   <span data-ttu-id="dd0eb-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-247">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-247">Requirements</span></span>

|<span data-ttu-id="dd0eb-248">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-248">Requirement</span></span>| <span data-ttu-id="dd0eb-249">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-251">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-251">1.0</span></span>|
|[<span data-ttu-id="dd0eb-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-253">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="dd0eb-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-256">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="dd0eb-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="dd0eb-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-261">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-262">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-262">Type</span></span>

*   [<span data-ttu-id="dd0eb-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd0eb-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-264">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-264">Requirements</span></span>

|<span data-ttu-id="dd0eb-265">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-265">Requirement</span></span>| <span data-ttu-id="dd0eb-266">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-268">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-268">1.0</span></span>|
|[<span data-ttu-id="dd0eb-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-270">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-272">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-273">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="dd0eb-274">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-274">internetMessageId :String</span></span>

<span data-ttu-id="dd0eb-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-277">Type</span><span class="sxs-lookup"><span data-stu-id="dd0eb-277">Type</span></span>

*   <span data-ttu-id="dd0eb-278">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-279">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-279">Requirements</span></span>

|<span data-ttu-id="dd0eb-280">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-280">Requirement</span></span>| <span data-ttu-id="dd0eb-281">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-283">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-283">1.0</span></span>|
|[<span data-ttu-id="dd0eb-284">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-285">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-287">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-288">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="dd0eb-289">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-289">itemClass :String</span></span>

<span data-ttu-id="dd0eb-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="dd0eb-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="dd0eb-294">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-294">Type</span></span> | <span data-ttu-id="dd0eb-295">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-295">Description</span></span> | <span data-ttu-id="dd0eb-296">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="dd0eb-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="dd0eb-297">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="dd0eb-297">Appointment items</span></span> | <span data-ttu-id="dd0eb-298">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="dd0eb-299">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="dd0eb-299">Message items</span></span> | <span data-ttu-id="dd0eb-300">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="dd0eb-301">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス `IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-302">Type</span><span class="sxs-lookup"><span data-stu-id="dd0eb-302">Type</span></span>

*   <span data-ttu-id="dd0eb-303">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-304">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-304">Requirements</span></span>

|<span data-ttu-id="dd0eb-305">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-305">Requirement</span></span>| <span data-ttu-id="dd0eb-306">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-307">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-308">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-308">1.0</span></span>|
|[<span data-ttu-id="dd0eb-309">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-310">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-311">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-312">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-313">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="dd0eb-314">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-314">(nullable) itemId :String</span></span>

<span data-ttu-id="dd0eb-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-317">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="dd0eb-318">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="dd0eb-319">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="dd0eb-320">詳細は、「[Outlook アドインからの Outlook REST API の使用](/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="dd0eb-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-323">Type</span><span class="sxs-lookup"><span data-stu-id="dd0eb-323">Type</span></span>

*   <span data-ttu-id="dd0eb-324">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-325">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-325">Requirements</span></span>

|<span data-ttu-id="dd0eb-326">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-326">Requirement</span></span>| <span data-ttu-id="dd0eb-327">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-328">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-329">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-329">1.0</span></span>|
|[<span data-ttu-id="dd0eb-330">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-331">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-332">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-333">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-334">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-334">Example</span></span>

<span data-ttu-id="dd0eb-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="dd0eb-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-337">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="dd0eb-338">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="dd0eb-339">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-340">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-340">Type</span></span>

*   [<span data-ttu-id="dd0eb-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="dd0eb-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-342">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-342">Requirements</span></span>

|<span data-ttu-id="dd0eb-343">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-343">Requirement</span></span>| <span data-ttu-id="dd0eb-344">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-345">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-346">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-346">1.0</span></span>|
|[<span data-ttu-id="dd0eb-347">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-348">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-349">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-350">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-351">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="dd0eb-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-352">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="dd0eb-353">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-354">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-354">Read mode</span></span>

<span data-ttu-id="dd0eb-355">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-356">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-356">Compose mode</span></span>

<span data-ttu-id="dd0eb-357">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-358">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-358">Type</span></span>

*   <span data-ttu-id="dd0eb-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-360">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-360">Requirements</span></span>

|<span data-ttu-id="dd0eb-361">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-361">Requirement</span></span>| <span data-ttu-id="dd0eb-362">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-363">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-364">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-364">1.0</span></span>|
|[<span data-ttu-id="dd0eb-365">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-366">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-367">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-368">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="dd0eb-369">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-369">normalizedSubject :String</span></span>

<span data-ttu-id="dd0eb-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="dd0eb-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-374">Type</span><span class="sxs-lookup"><span data-stu-id="dd0eb-374">Type</span></span>

*   <span data-ttu-id="dd0eb-375">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-376">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-376">Requirements</span></span>

|<span data-ttu-id="dd0eb-377">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-377">Requirement</span></span>| <span data-ttu-id="dd0eb-378">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-379">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-380">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-380">1.0</span></span>|
|[<span data-ttu-id="dd0eb-381">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-382">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-383">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-384">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-385">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="dd0eb-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-386">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="dd0eb-387">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-388">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-388">Type</span></span>

*   [<span data-ttu-id="dd0eb-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="dd0eb-389">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-390">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-390">Requirements</span></span>

|<span data-ttu-id="dd0eb-391">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-391">Requirement</span></span>| <span data-ttu-id="dd0eb-392">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-393">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-394">1.3</span><span class="sxs-lookup"><span data-stu-id="dd0eb-394">1.3</span></span>|
|[<span data-ttu-id="dd0eb-395">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-396">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-397">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-398">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-399">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="dd0eb-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-400">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="dd0eb-401">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="dd0eb-402">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-403">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-403">Read mode</span></span>

<span data-ttu-id="dd0eb-404">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-405">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-405">Compose mode</span></span>

<span data-ttu-id="dd0eb-406">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-407">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-407">Type</span></span>

*   <span data-ttu-id="dd0eb-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-409">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-409">Requirements</span></span>

|<span data-ttu-id="dd0eb-410">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-410">Requirement</span></span>| <span data-ttu-id="dd0eb-411">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-412">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-413">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-413">1.0</span></span>|
|[<span data-ttu-id="dd0eb-414">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-415">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-416">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-417">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="dd0eb-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-418">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="dd0eb-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-421">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-421">Type</span></span>

*   [<span data-ttu-id="dd0eb-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd0eb-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-423">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-423">Requirements</span></span>

|<span data-ttu-id="dd0eb-424">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-424">Requirement</span></span>| <span data-ttu-id="dd0eb-425">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-426">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-427">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-427">1.0</span></span>|
|[<span data-ttu-id="dd0eb-428">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-429">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-430">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-431">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-432">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="dd0eb-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-433">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="dd0eb-434">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="dd0eb-435">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-436">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-436">Read mode</span></span>

<span data-ttu-id="dd0eb-437">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-438">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-438">Compose mode</span></span>

<span data-ttu-id="dd0eb-439">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-440">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-440">Type</span></span>

*   <span data-ttu-id="dd0eb-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-442">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-442">Requirements</span></span>

|<span data-ttu-id="dd0eb-443">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-443">Requirement</span></span>| <span data-ttu-id="dd0eb-444">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-445">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-446">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-446">1.0</span></span>|
|[<span data-ttu-id="dd0eb-447">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-448">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-449">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-450">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="dd0eb-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-451">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="dd0eb-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="dd0eb-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-456">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0eb-457">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-457">Type</span></span>

*   [<span data-ttu-id="dd0eb-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="dd0eb-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="dd0eb-459">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-459">Requirements</span></span>

|<span data-ttu-id="dd0eb-460">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-460">Requirement</span></span>| <span data-ttu-id="dd0eb-461">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-462">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-463">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-463">1.0</span></span>|
|[<span data-ttu-id="dd0eb-464">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-465">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-466">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-467">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-468">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="dd0eb-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-469">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="dd0eb-470">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="dd0eb-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-473">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-473">Read mode</span></span>

<span data-ttu-id="dd0eb-474">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-475">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-475">Compose mode</span></span>

<span data-ttu-id="dd0eb-476">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="dd0eb-477">[`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="dd0eb-478">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="dd0eb-479">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-479">Type</span></span>

*   <span data-ttu-id="dd0eb-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-481">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-481">Requirements</span></span>

|<span data-ttu-id="dd0eb-482">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-482">Requirement</span></span>| <span data-ttu-id="dd0eb-483">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-484">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-485">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-485">1.0</span></span>|
|[<span data-ttu-id="dd0eb-486">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-487">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-488">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-489">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-489">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="dd0eb-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="dd0eb-491">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="dd0eb-492">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-493">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-493">Read mode</span></span>

<span data-ttu-id="dd0eb-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-496">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-496">Compose mode</span></span>

<span data-ttu-id="dd0eb-497">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-498">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-498">Type</span></span>

*   <span data-ttu-id="dd0eb-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-500">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-500">Requirements</span></span>

|<span data-ttu-id="dd0eb-501">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-501">Requirement</span></span>| <span data-ttu-id="dd0eb-502">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-504">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-504">1.0</span></span>|
|[<span data-ttu-id="dd0eb-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-506">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-508">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-508">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="dd0eb-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="dd0eb-510">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="dd0eb-511">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="dd0eb-512">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-512">Read mode</span></span>

<span data-ttu-id="dd0eb-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="dd0eb-515">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-515">Compose mode</span></span>

<span data-ttu-id="dd0eb-516">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="dd0eb-517">型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-517">Type</span></span>

*   <span data-ttu-id="dd0eb-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-519">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-519">Requirements</span></span>

|<span data-ttu-id="dd0eb-520">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-520">Requirement</span></span>| <span data-ttu-id="dd0eb-521">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-523">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-523">1.0</span></span>|
|[<span data-ttu-id="dd0eb-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-525">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-527">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="dd0eb-528">メソッド</span><span class="sxs-lookup"><span data-stu-id="dd0eb-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="dd0eb-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="dd0eb-530">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="dd0eb-531">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="dd0eb-532">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-533">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-533">Parameters</span></span>

|<span data-ttu-id="dd0eb-534">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-534">Name</span></span>| <span data-ttu-id="dd0eb-535">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-535">Type</span></span>| <span data-ttu-id="dd0eb-536">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-536">Attributes</span></span>| <span data-ttu-id="dd0eb-537">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="dd0eb-538">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-538">String</span></span>||<span data-ttu-id="dd0eb-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="dd0eb-541">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-541">String</span></span>||<span data-ttu-id="dd0eb-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="dd0eb-544">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-544">Object</span></span>| <span data-ttu-id="dd0eb-545">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-545">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-546">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-547">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-547">Object</span></span>| <span data-ttu-id="dd0eb-548">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-548">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-549">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd0eb-550">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-550">function</span></span>| <span data-ttu-id="dd0eb-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-551">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-552">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd0eb-553">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="dd0eb-554">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd0eb-555">エラー</span><span class="sxs-lookup"><span data-stu-id="dd0eb-555">Errors</span></span>

| <span data-ttu-id="dd0eb-556">エラー コード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-556">Error code</span></span> | <span data-ttu-id="dd0eb-557">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="dd0eb-558">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="dd0eb-559">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="dd0eb-560">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-561">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-561">Requirements</span></span>

|<span data-ttu-id="dd0eb-562">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-562">Requirement</span></span>| <span data-ttu-id="dd0eb-563">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-565">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0eb-565">1.1</span></span>|
|[<span data-ttu-id="dd0eb-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-569">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-570">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="dd0eb-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="dd0eb-572">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="dd0eb-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="dd0eb-576">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="dd0eb-577">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-578">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-578">Parameters</span></span>

|<span data-ttu-id="dd0eb-579">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-579">Name</span></span>| <span data-ttu-id="dd0eb-580">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-580">Type</span></span>| <span data-ttu-id="dd0eb-581">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-581">Attributes</span></span>| <span data-ttu-id="dd0eb-582">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="dd0eb-583">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-583">String</span></span>||<span data-ttu-id="dd0eb-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="dd0eb-586">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-586">String</span></span>||<span data-ttu-id="dd0eb-587">添付するアイテムの件名。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-587">The subject of the item to be attached.</span></span> <span data-ttu-id="dd0eb-588">最大の長さは、255 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="dd0eb-589">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-589">Object</span></span>| <span data-ttu-id="dd0eb-590">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-590">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-591">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-592">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dd0eb-592">Object</span></span>| <span data-ttu-id="dd0eb-593">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-593">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-594">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd0eb-595">関数</span><span class="sxs-lookup"><span data-stu-id="dd0eb-595">function</span></span>| <span data-ttu-id="dd0eb-596">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-596">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd0eb-598">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="dd0eb-599">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd0eb-600">エラー</span><span class="sxs-lookup"><span data-stu-id="dd0eb-600">Errors</span></span>

| <span data-ttu-id="dd0eb-601">エラー コード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-601">Error code</span></span> | <span data-ttu-id="dd0eb-602">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="dd0eb-603">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-604">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-604">Requirements</span></span>

|<span data-ttu-id="dd0eb-605">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-605">Requirement</span></span>| <span data-ttu-id="dd0eb-606">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-608">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0eb-608">1.1</span></span>|
|[<span data-ttu-id="dd0eb-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-612">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-613">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-613">Example</span></span>

<span data-ttu-id="dd0eb-614">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="dd0eb-615">close()</span><span class="sxs-lookup"><span data-stu-id="dd0eb-615">close()</span></span>

<span data-ttu-id="dd0eb-616">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="dd0eb-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-619">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="dd0eb-620">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-621">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-621">Requirements</span></span>

|<span data-ttu-id="dd0eb-622">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-622">Requirement</span></span>| <span data-ttu-id="dd0eb-623">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-624">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-625">1.3</span><span class="sxs-lookup"><span data-stu-id="dd0eb-625">1.3</span></span>|
|[<span data-ttu-id="dd0eb-626">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-627">制限あり</span><span class="sxs-lookup"><span data-stu-id="dd0eb-627">Restricted</span></span>|
|[<span data-ttu-id="dd0eb-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-629">新規作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="dd0eb-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="dd0eb-631">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-632">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dd0eb-633">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="dd0eb-634">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="dd0eb-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-638">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-638">Parameters</span></span>

|<span data-ttu-id="dd0eb-639">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-639">Name</span></span>| <span data-ttu-id="dd0eb-640">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-640">Type</span></span>| <span data-ttu-id="dd0eb-641">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="dd0eb-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-642">String &#124; Object</span></span>| |<span data-ttu-id="dd0eb-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="dd0eb-645">**または**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-645">**OR**</span></span><br/><span data-ttu-id="dd0eb-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="dd0eb-648">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-648">String</span></span> | <span data-ttu-id="dd0eb-649">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-649">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="dd0eb-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dd0eb-653">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-653">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-654">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="dd0eb-655">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-655">String</span></span> | | <span data-ttu-id="dd0eb-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="dd0eb-658">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-658">String</span></span> | | <span data-ttu-id="dd0eb-659">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="dd0eb-660">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0eb-660">String</span></span> | | <span data-ttu-id="dd0eb-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="dd0eb-663">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-663">String</span></span> | | <span data-ttu-id="dd0eb-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="dd0eb-667">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-667">function</span></span> | <span data-ttu-id="dd0eb-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-668">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-669">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-670">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-670">Requirements</span></span>

|<span data-ttu-id="dd0eb-671">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-671">Requirement</span></span>| <span data-ttu-id="dd0eb-672">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-673">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-674">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-674">1.0</span></span>|
|[<span data-ttu-id="dd0eb-675">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-676">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-678">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd0eb-679">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-679">Examples</span></span>

<span data-ttu-id="dd0eb-680">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="dd0eb-681">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="dd0eb-682">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="dd0eb-683">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="dd0eb-684">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="dd0eb-685">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="dd0eb-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="dd0eb-687">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-688">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dd0eb-689">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="dd0eb-690">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="dd0eb-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-694">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-694">Parameters</span></span>

|<span data-ttu-id="dd0eb-695">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-695">Name</span></span>| <span data-ttu-id="dd0eb-696">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-696">Type</span></span>| <span data-ttu-id="dd0eb-697">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="dd0eb-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-698">String &#124; Object</span></span>| | <span data-ttu-id="dd0eb-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="dd0eb-701">**または**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-701">**OR**</span></span><br/><span data-ttu-id="dd0eb-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="dd0eb-704">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-704">String</span></span> | <span data-ttu-id="dd0eb-705">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-705">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="dd0eb-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dd0eb-709">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-709">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-710">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="dd0eb-711">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-711">String</span></span> | | <span data-ttu-id="dd0eb-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="dd0eb-714">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-714">String</span></span> | | <span data-ttu-id="dd0eb-715">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="dd0eb-716">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0eb-716">String</span></span> | | <span data-ttu-id="dd0eb-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="dd0eb-719">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-719">String</span></span> | | <span data-ttu-id="dd0eb-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="dd0eb-723">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-723">function</span></span> | <span data-ttu-id="dd0eb-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-724">&lt;optional&gt;</span></span> | <span data-ttu-id="dd0eb-725">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-726">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-726">Requirements</span></span>

|<span data-ttu-id="dd0eb-727">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-727">Requirement</span></span>| <span data-ttu-id="dd0eb-728">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-729">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-730">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-730">1.0</span></span>|
|[<span data-ttu-id="dd0eb-731">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-732">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-734">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd0eb-735">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-735">Examples</span></span>

<span data-ttu-id="dd0eb-736">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="dd0eb-737">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="dd0eb-738">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="dd0eb-739">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="dd0eb-740">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="dd0eb-741">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="dd0eb-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="dd0eb-743">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-744">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-745">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-745">Requirements</span></span>

|<span data-ttu-id="dd0eb-746">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-746">Requirement</span></span>| <span data-ttu-id="dd0eb-747">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-749">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-749">1.0</span></span>|
|[<span data-ttu-id="dd0eb-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-751">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-753">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-754">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-754">Returns:</span></span>

<span data-ttu-id="dd0eb-755">型:[Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="dd0eb-756">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-756">Example</span></span>

<span data-ttu-id="dd0eb-757">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="dd0eb-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="dd0eb-759">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-760">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-761">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-761">Parameters</span></span>

|<span data-ttu-id="dd0eb-762">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-762">Name</span></span>| <span data-ttu-id="dd0eb-763">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-763">Type</span></span>| <span data-ttu-id="dd0eb-764">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="dd0eb-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="dd0eb-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="dd0eb-766">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-767">Requirements</span><span class="sxs-lookup"><span data-stu-id="dd0eb-767">Requirements</span></span>

|<span data-ttu-id="dd0eb-768">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-768">Requirement</span></span>| <span data-ttu-id="dd0eb-769">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-770">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-771">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-771">1.0</span></span>|
|[<span data-ttu-id="dd0eb-772">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-773">制限あり</span><span class="sxs-lookup"><span data-stu-id="dd0eb-773">Restricted</span></span>|
|[<span data-ttu-id="dd0eb-774">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-775">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-776">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-776">Returns:</span></span>

<span data-ttu-id="dd0eb-777">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="dd0eb-778">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="dd0eb-779">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="dd0eb-780">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="dd0eb-781">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-781">Value of `entityType`</span></span> | <span data-ttu-id="dd0eb-782">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="dd0eb-782">Type of objects in returned array</span></span> | <span data-ttu-id="dd0eb-783">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="dd0eb-784">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0eb-784">String</span></span> | <span data-ttu-id="dd0eb-785">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="dd0eb-786">連絡先</span><span class="sxs-lookup"><span data-stu-id="dd0eb-786">Contact</span></span> | <span data-ttu-id="dd0eb-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="dd0eb-788">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0eb-788">String</span></span> | <span data-ttu-id="dd0eb-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="dd0eb-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="dd0eb-790">MeetingSuggestion</span></span> | <span data-ttu-id="dd0eb-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="dd0eb-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="dd0eb-792">PhoneNumber</span></span> | <span data-ttu-id="dd0eb-793">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="dd0eb-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="dd0eb-794">TaskSuggestion</span></span> | <span data-ttu-id="dd0eb-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="dd0eb-796">文字列</span><span class="sxs-lookup"><span data-stu-id="dd0eb-796">String</span></span> | <span data-ttu-id="dd0eb-797">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="dd0eb-797">**Restricted**</span></span> |

<span data-ttu-id="dd0eb-798">型:Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="dd0eb-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="dd0eb-799">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-799">Example</span></span>

<span data-ttu-id="dd0eb-800">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="dd0eb-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="dd0eb-802">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-803">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dd0eb-804">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-805">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-805">Parameters</span></span>

|<span data-ttu-id="dd0eb-806">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-806">Name</span></span>| <span data-ttu-id="dd0eb-807">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-807">Type</span></span>| <span data-ttu-id="dd0eb-808">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="dd0eb-809">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-809">String</span></span>|<span data-ttu-id="dd0eb-810">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-811">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-811">Requirements</span></span>

|<span data-ttu-id="dd0eb-812">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-812">Requirement</span></span>| <span data-ttu-id="dd0eb-813">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-814">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-815">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-815">1.0</span></span>|
|[<span data-ttu-id="dd0eb-816">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-817">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-819">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-820">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-820">Returns:</span></span>

<span data-ttu-id="dd0eb-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="dd0eb-823">型:Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="dd0eb-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="dd0eb-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="dd0eb-825">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-826">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dd0eb-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="dd0eb-830">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="dd0eb-831">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="dd0eb-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0eb-835">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-835">Requirements</span></span>

|<span data-ttu-id="dd0eb-836">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-836">Requirement</span></span>| <span data-ttu-id="dd0eb-837">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-838">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-839">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-839">1.0</span></span>|
|[<span data-ttu-id="dd0eb-840">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-841">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-842">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-843">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-844">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-844">Returns:</span></span>

<span data-ttu-id="dd0eb-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="dd0eb-847">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="dd0eb-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="dd0eb-848">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="dd0eb-849">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-849">Example</span></span>

<span data-ttu-id="dd0eb-850">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="dd0eb-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="dd0eb-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="dd0eb-852">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-853">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="dd0eb-854">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="dd0eb-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-857">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-857">Parameters</span></span>

|<span data-ttu-id="dd0eb-858">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-858">Name</span></span>| <span data-ttu-id="dd0eb-859">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-859">Type</span></span>| <span data-ttu-id="dd0eb-860">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="dd0eb-861">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-861">String</span></span>|<span data-ttu-id="dd0eb-862">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-863">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-863">Requirements</span></span>

|<span data-ttu-id="dd0eb-864">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-864">Requirement</span></span>| <span data-ttu-id="dd0eb-865">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-867">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-867">1.0</span></span>|
|[<span data-ttu-id="dd0eb-868">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-869">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-871">読み取り</span><span class="sxs-lookup"><span data-stu-id="dd0eb-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-872">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-872">Returns:</span></span>

<span data-ttu-id="dd0eb-873">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="dd0eb-874">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="dd0eb-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="dd0eb-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="dd0eb-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="dd0eb-876">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="dd0eb-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="dd0eb-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="dd0eb-878">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="dd0eb-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-881">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-881">Parameters</span></span>

|<span data-ttu-id="dd0eb-882">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-882">Name</span></span>| <span data-ttu-id="dd0eb-883">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-883">Type</span></span>| <span data-ttu-id="dd0eb-884">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-884">Attributes</span></span>| <span data-ttu-id="dd0eb-885">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="dd0eb-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd0eb-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="dd0eb-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="dd0eb-890">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-890">Object</span></span>| <span data-ttu-id="dd0eb-891">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-891">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-892">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-893">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dd0eb-893">Object</span></span>| <span data-ttu-id="dd0eb-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-894">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-895">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd0eb-896">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-896">function</span></span>||<span data-ttu-id="dd0eb-897">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd0eb-898">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="dd0eb-899">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-900">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-900">Requirements</span></span>

|<span data-ttu-id="dd0eb-901">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-901">Requirement</span></span>| <span data-ttu-id="dd0eb-902">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-903">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-904">1.2</span><span class="sxs-lookup"><span data-stu-id="dd0eb-904">1.2</span></span>|
|[<span data-ttu-id="dd0eb-905">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-908">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd0eb-909">戻り値:</span><span class="sxs-lookup"><span data-stu-id="dd0eb-909">Returns:</span></span>

<span data-ttu-id="dd0eb-910">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="dd0eb-911">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="dd0eb-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="dd0eb-912">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="dd0eb-913">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="dd0eb-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="dd0eb-915">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="dd0eb-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-919">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-919">Parameters</span></span>

|<span data-ttu-id="dd0eb-920">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-920">Name</span></span>| <span data-ttu-id="dd0eb-921">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-921">Type</span></span>| <span data-ttu-id="dd0eb-922">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-922">Attributes</span></span>| <span data-ttu-id="dd0eb-923">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dd0eb-924">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-924">function</span></span>||<span data-ttu-id="dd0eb-925">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd0eb-926">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="dd0eb-927">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="dd0eb-928">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-928">Object</span></span>| <span data-ttu-id="dd0eb-929">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-929">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-930">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="dd0eb-931">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-932">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-932">Requirements</span></span>

|<span data-ttu-id="dd0eb-933">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-933">Requirement</span></span>| <span data-ttu-id="dd0eb-934">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-936">1.0</span><span class="sxs-lookup"><span data-stu-id="dd0eb-936">1.0</span></span>|
|[<span data-ttu-id="dd0eb-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-938">ReadItem</span></span>|
|[<span data-ttu-id="dd0eb-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-940">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dd0eb-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-941">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-941">Example</span></span>

<span data-ttu-id="dd0eb-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="dd0eb-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd0eb-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="dd0eb-946">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="dd0eb-p165">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-951">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-951">Parameters</span></span>

|<span data-ttu-id="dd0eb-952">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-952">Name</span></span>| <span data-ttu-id="dd0eb-953">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-953">Type</span></span>| <span data-ttu-id="dd0eb-954">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-954">Attributes</span></span>| <span data-ttu-id="dd0eb-955">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="dd0eb-956">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-956">String</span></span>||<span data-ttu-id="dd0eb-957">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="dd0eb-958">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dd0eb-958">Object</span></span>| <span data-ttu-id="dd0eb-959">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-959">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-960">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-961">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-961">Object</span></span>| <span data-ttu-id="dd0eb-962">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-962">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-963">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="dd0eb-964">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-964">function</span></span>| <span data-ttu-id="dd0eb-965">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-965">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-966">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="dd0eb-967">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd0eb-968">エラー</span><span class="sxs-lookup"><span data-stu-id="dd0eb-968">Errors</span></span>

| <span data-ttu-id="dd0eb-969">エラー コード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-969">Error code</span></span> | <span data-ttu-id="dd0eb-970">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="dd0eb-971">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-972">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-972">Requirements</span></span>

|<span data-ttu-id="dd0eb-973">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-973">Requirement</span></span>| <span data-ttu-id="dd0eb-974">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-975">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-976">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0eb-976">1.1</span></span>|
|[<span data-ttu-id="dd0eb-977">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-979">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-980">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-981">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-981">Example</span></span>

<span data-ttu-id="dd0eb-982">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="dd0eb-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="dd0eb-984">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="dd0eb-p166">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-988">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="dd0eb-989">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="dd0eb-p168">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="dd0eb-993">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="dd0eb-994">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-994">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="dd0eb-995">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-995">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="dd0eb-996">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-996">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-997">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-997">Parameters</span></span>

|<span data-ttu-id="dd0eb-998">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-998">Name</span></span>| <span data-ttu-id="dd0eb-999">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-999">Type</span></span>| <span data-ttu-id="dd0eb-1000">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1000">Attributes</span></span>| <span data-ttu-id="dd0eb-1001">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1001">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="dd0eb-1002">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1002">Object</span></span>| <span data-ttu-id="dd0eb-1003">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1003">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-1004">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1004">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-1005">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1005">Object</span></span>| <span data-ttu-id="dd0eb-1006">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-1007">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1007">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="dd0eb-1008">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1008">function</span></span>||<span data-ttu-id="dd0eb-1009">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd0eb-1010">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1010">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0eb-1011">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1011">Requirements</span></span>

|<span data-ttu-id="dd0eb-1012">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1012">Requirement</span></span>| <span data-ttu-id="dd0eb-1013">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-1014">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-1015">1.3</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1015">1.3</span></span>|
|[<span data-ttu-id="dd0eb-1016">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-1018">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-1019">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1019">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="dd0eb-1020">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1020">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="dd0eb-p170">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="dd0eb-1023">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1023">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="dd0eb-1024">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1024">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="dd0eb-p171">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd0eb-1028">パラメーター</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1028">Parameters</span></span>

|<span data-ttu-id="dd0eb-1029">名前</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1029">Name</span></span>| <span data-ttu-id="dd0eb-1030">種類</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1030">Type</span></span>| <span data-ttu-id="dd0eb-1031">属性</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1031">Attributes</span></span>| <span data-ttu-id="dd0eb-1032">説明</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1032">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dd0eb-1033">String</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1033">String</span></span>||<span data-ttu-id="dd0eb-p172">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="dd0eb-1037">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1037">Object</span></span>| <span data-ttu-id="dd0eb-1038">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-1039">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1039">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="dd0eb-1040">Object</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1040">Object</span></span>| <span data-ttu-id="dd0eb-1041">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-1042">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1042">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="dd0eb-1043">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1043">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="dd0eb-1044">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1044">&lt;optional&gt;</span></span>|<span data-ttu-id="dd0eb-p173">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="dd0eb-p174">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="dd0eb-1049">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1049">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="dd0eb-1050">function</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1050">function</span></span>||<span data-ttu-id="dd0eb-1051">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1051">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0eb-1052">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1052">Requirements</span></span>

|<span data-ttu-id="dd0eb-1053">要件</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1053">Requirement</span></span>| <span data-ttu-id="dd0eb-1054">値</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1054">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0eb-1055">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1055">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd0eb-1056">1.2</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1056">1.2</span></span>|
|[<span data-ttu-id="dd0eb-1057">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1057">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd0eb-1058">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1058">ReadWriteItem</span></span>|
|[<span data-ttu-id="dd0eb-1059">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1059">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd0eb-1060">作成</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1060">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="dd0eb-1061">例</span><span class="sxs-lookup"><span data-stu-id="dd0eb-1061">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
