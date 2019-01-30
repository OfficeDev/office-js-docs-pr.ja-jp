---
title: Office.context.mailbox.item ・要件設定 1.3
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: c418c69e369e5f8ed6da151345013897f1a87e26
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387325"
---
# <a name="item"></a><span data-ttu-id="c940b-102">item</span><span class="sxs-lookup"><span data-stu-id="c940b-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c940b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c940b-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c940b-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-106">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-106">Requirements</span></span>

|<span data-ttu-id="c940b-107">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-107">Requirement</span></span>| <span data-ttu-id="c940b-108">値</span><span class="sxs-lookup"><span data-stu-id="c940b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-110">1.0</span></span>|
|[<span data-ttu-id="c940b-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="c940b-112">Restricted</span></span>|
|[<span data-ttu-id="c940b-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="c940b-115">例</span><span class="sxs-lookup"><span data-stu-id="c940b-115">Example</span></span>

<span data-ttu-id="c940b-116">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="c940b-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="c940b-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="c940b-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="c940b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c940b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="c940b-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-121">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="c940b-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c940b-122">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c940b-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-123">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-123">Type:</span></span>

*   <span data-ttu-id="c940b-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c940b-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-125">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-125">Requirements</span></span>

|<span data-ttu-id="c940b-126">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-126">Requirement</span></span>| <span data-ttu-id="c940b-127">値</span><span class="sxs-lookup"><span data-stu-id="c940b-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-128">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-129">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-129">1.0</span></span>|
|[<span data-ttu-id="c940b-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-131">ReadItem</span></span>|
|[<span data-ttu-id="c940b-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-134">例</span><span class="sxs-lookup"><span data-stu-id="c940b-134">Example</span></span>

<span data-ttu-id="c940b-135">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="c940b-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c940b-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-136">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c940b-137">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c940b-138">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-139">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-139">Type:</span></span>

*   [<span data-ttu-id="c940b-140">Recipients</span><span class="sxs-lookup"><span data-stu-id="c940b-140">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c940b-141">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-141">Requirements</span></span>

|<span data-ttu-id="c940b-142">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-142">Requirement</span></span>| <span data-ttu-id="c940b-143">値</span><span class="sxs-lookup"><span data-stu-id="c940b-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-145">1.1</span><span class="sxs-lookup"><span data-stu-id="c940b-145">1.1</span></span>|
|[<span data-ttu-id="c940b-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-147">ReadItem</span></span>|
|[<span data-ttu-id="c940b-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-149">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-150">例</span><span class="sxs-lookup"><span data-stu-id="c940b-150">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="c940b-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="c940b-151">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="c940b-152">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-153">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-153">Type:</span></span>

*   [<span data-ttu-id="c940b-154">Body</span><span class="sxs-lookup"><span data-stu-id="c940b-154">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="c940b-155">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-155">Requirements</span></span>

|<span data-ttu-id="c940b-156">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-156">Requirement</span></span>| <span data-ttu-id="c940b-157">値</span><span class="sxs-lookup"><span data-stu-id="c940b-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c940b-159">1.1</span></span>|
|[<span data-ttu-id="c940b-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-161">ReadItem</span></span>|
|[<span data-ttu-id="c940b-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c940b-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c940b-165">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c940b-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c940b-166">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-167">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-167">Read mode</span></span>

<span data-ttu-id="c940b-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-170">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-170">Compose mode</span></span>

<span data-ttu-id="c940b-171">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-172">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-172">Type:</span></span>

*   <span data-ttu-id="c940b-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-174">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-174">Requirements</span></span>

|<span data-ttu-id="c940b-175">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-175">Requirement</span></span>| <span data-ttu-id="c940b-176">値</span><span class="sxs-lookup"><span data-stu-id="c940b-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-177">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-178">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-178">1.0</span></span>|
|[<span data-ttu-id="c940b-179">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-180">ReadItem</span></span>|
|[<span data-ttu-id="c940b-181">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-182">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-183">例</span><span class="sxs-lookup"><span data-stu-id="c940b-183">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c940b-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c940b-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="c940b-185">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c940b-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="c940b-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c940b-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-190">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-190">Type:</span></span>

*   <span data-ttu-id="c940b-191">String</span><span class="sxs-lookup"><span data-stu-id="c940b-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-192">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-192">Requirements</span></span>

|<span data-ttu-id="c940b-193">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-193">Requirement</span></span>| <span data-ttu-id="c940b-194">値</span><span class="sxs-lookup"><span data-stu-id="c940b-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-195">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-196">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-196">1.0</span></span>|
|[<span data-ttu-id="c940b-197">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-198">ReadItem</span></span>|
|[<span data-ttu-id="c940b-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-200">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c940b-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c940b-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="c940b-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-204">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-204">Type:</span></span>

*   <span data-ttu-id="c940b-205">日付</span><span class="sxs-lookup"><span data-stu-id="c940b-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-206">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-206">Requirements</span></span>

|<span data-ttu-id="c940b-207">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-207">Requirement</span></span>| <span data-ttu-id="c940b-208">値</span><span class="sxs-lookup"><span data-stu-id="c940b-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-210">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-210">1.0</span></span>|
|[<span data-ttu-id="c940b-211">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-212">ReadItem</span></span>|
|[<span data-ttu-id="c940b-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-214">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-215">例</span><span class="sxs-lookup"><span data-stu-id="c940b-215">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c940b-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c940b-216">dateTimeModified :Date</span></span>

<span data-ttu-id="c940b-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-219">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-220">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-220">Type:</span></span>

*   <span data-ttu-id="c940b-221">日付</span><span class="sxs-lookup"><span data-stu-id="c940b-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-222">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-222">Requirements</span></span>

|<span data-ttu-id="c940b-223">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-223">Requirement</span></span>| <span data-ttu-id="c940b-224">値</span><span class="sxs-lookup"><span data-stu-id="c940b-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-226">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-226">1.0</span></span>|
|[<span data-ttu-id="c940b-227">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-228">ReadItem</span></span>|
|[<span data-ttu-id="c940b-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-230">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-231">例</span><span class="sxs-lookup"><span data-stu-id="c940b-231">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="c940b-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c940b-232">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="c940b-233">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c940b-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-236">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-236">Read mode</span></span>

<span data-ttu-id="c940b-237">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-238">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-238">Compose mode</span></span>

<span data-ttu-id="c940b-239">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c940b-240">[`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c940b-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-241">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-241">Type:</span></span>

*   <span data-ttu-id="c940b-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c940b-242">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-243">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-243">Requirements</span></span>

|<span data-ttu-id="c940b-244">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-244">Requirement</span></span>| <span data-ttu-id="c940b-245">値</span><span class="sxs-lookup"><span data-stu-id="c940b-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-246">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-247">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-247">1.0</span></span>|
|[<span data-ttu-id="c940b-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-249">ReadItem</span></span>|
|[<span data-ttu-id="c940b-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-251">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-252">例</span><span class="sxs-lookup"><span data-stu-id="c940b-252">Example</span></span>

<span data-ttu-id="c940b-253">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c940b-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c940b-254">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c940b-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c940b-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-259">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c940b-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-260">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-260">Type:</span></span>

*   [<span data-ttu-id="c940b-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c940b-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c940b-262">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-262">Requirements</span></span>

|<span data-ttu-id="c940b-263">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-263">Requirement</span></span>| <span data-ttu-id="c940b-264">値</span><span class="sxs-lookup"><span data-stu-id="c940b-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-266">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-266">1.0</span></span>|
|[<span data-ttu-id="c940b-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-268">ReadItem</span></span>|
|[<span data-ttu-id="c940b-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-270">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c940b-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c940b-271">internetMessageId :String</span></span>

<span data-ttu-id="c940b-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-274">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-274">Type:</span></span>

*   <span data-ttu-id="c940b-275">String</span><span class="sxs-lookup"><span data-stu-id="c940b-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-276">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-276">Requirements</span></span>

|<span data-ttu-id="c940b-277">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-277">Requirement</span></span>| <span data-ttu-id="c940b-278">値</span><span class="sxs-lookup"><span data-stu-id="c940b-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-280">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-280">1.0</span></span>|
|[<span data-ttu-id="c940b-281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-282">ReadItem</span></span>|
|[<span data-ttu-id="c940b-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-284">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-285">例</span><span class="sxs-lookup"><span data-stu-id="c940b-285">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c940b-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c940b-286">itemClass :String</span></span>

<span data-ttu-id="c940b-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c940b-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c940b-291">型</span><span class="sxs-lookup"><span data-stu-id="c940b-291">Type</span></span> | <span data-ttu-id="c940b-292">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-292">Description</span></span> | <span data-ttu-id="c940b-293">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="c940b-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c940b-294">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="c940b-294">Appointment items</span></span> | <span data-ttu-id="c940b-295">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c940b-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="c940b-296">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="c940b-296">Message items</span></span> | <span data-ttu-id="c940b-297">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c940b-298">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-299">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-299">Type:</span></span>

*   <span data-ttu-id="c940b-300">String</span><span class="sxs-lookup"><span data-stu-id="c940b-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-301">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-301">Requirements</span></span>

|<span data-ttu-id="c940b-302">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-302">Requirement</span></span>| <span data-ttu-id="c940b-303">値</span><span class="sxs-lookup"><span data-stu-id="c940b-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-305">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-305">1.0</span></span>|
|[<span data-ttu-id="c940b-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-307">ReadItem</span></span>|
|[<span data-ttu-id="c940b-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-310">例</span><span class="sxs-lookup"><span data-stu-id="c940b-310">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c940b-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c940b-311">(nullable) itemId :String</span></span>

<span data-ttu-id="c940b-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-314">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="c940b-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c940b-315">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="c940b-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c940b-316">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c940b-316">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c940b-317">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c940b-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c940b-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-320">種類:</span><span class="sxs-lookup"><span data-stu-id="c940b-320">Type:</span></span>

*   <span data-ttu-id="c940b-321">String</span><span class="sxs-lookup"><span data-stu-id="c940b-321">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-322">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-322">Requirements</span></span>

|<span data-ttu-id="c940b-323">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-323">Requirement</span></span>| <span data-ttu-id="c940b-324">値</span><span class="sxs-lookup"><span data-stu-id="c940b-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-325">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-326">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-326">1.0</span></span>|
|[<span data-ttu-id="c940b-327">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-328">ReadItem</span></span>|
|[<span data-ttu-id="c940b-329">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-330">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-330">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-331">例</span><span class="sxs-lookup"><span data-stu-id="c940b-331">Example</span></span>

<span data-ttu-id="c940b-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="c940b-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c940b-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c940b-335">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-335">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c940b-336">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="c940b-336">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-337">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-337">Type:</span></span>

*   [<span data-ttu-id="c940b-338">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c940b-338">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c940b-339">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-339">Requirements</span></span>

|<span data-ttu-id="c940b-340">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-340">Requirement</span></span>| <span data-ttu-id="c940b-341">値</span><span class="sxs-lookup"><span data-stu-id="c940b-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-342">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-343">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-343">1.0</span></span>|
|[<span data-ttu-id="c940b-344">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-345">ReadItem</span></span>|
|[<span data-ttu-id="c940b-346">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-347">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-348">例</span><span class="sxs-lookup"><span data-stu-id="c940b-348">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="c940b-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="c940b-349">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="c940b-350">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-350">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-351">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-351">Read mode</span></span>

<span data-ttu-id="c940b-352">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-352">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-353">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-353">Compose mode</span></span>

<span data-ttu-id="c940b-354">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-354">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-355">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-355">Type:</span></span>

*   <span data-ttu-id="c940b-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="c940b-356">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-357">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-357">Requirements</span></span>

|<span data-ttu-id="c940b-358">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-358">Requirement</span></span>| <span data-ttu-id="c940b-359">値</span><span class="sxs-lookup"><span data-stu-id="c940b-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-360">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-361">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-361">1.0</span></span>|
|[<span data-ttu-id="c940b-362">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-363">ReadItem</span></span>|
|[<span data-ttu-id="c940b-364">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-365">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-366">例</span><span class="sxs-lookup"><span data-stu-id="c940b-366">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c940b-367">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c940b-367">normalizedSubject :String</span></span>

<span data-ttu-id="c940b-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c940b-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-372">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-372">Type:</span></span>

*   <span data-ttu-id="c940b-373">String</span><span class="sxs-lookup"><span data-stu-id="c940b-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-374">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-374">Requirements</span></span>

|<span data-ttu-id="c940b-375">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-375">Requirement</span></span>| <span data-ttu-id="c940b-376">値</span><span class="sxs-lookup"><span data-stu-id="c940b-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-377">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-378">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-378">1.0</span></span>|
|[<span data-ttu-id="c940b-379">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-380">ReadItem</span></span>|
|[<span data-ttu-id="c940b-381">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-382">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-383">例</span><span class="sxs-lookup"><span data-stu-id="c940b-383">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="c940b-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c940b-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="c940b-385">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-385">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-386">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-386">Type:</span></span>

*   [<span data-ttu-id="c940b-387">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c940b-387">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c940b-388">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-388">Requirements</span></span>

|<span data-ttu-id="c940b-389">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-389">Requirement</span></span>| <span data-ttu-id="c940b-390">値</span><span class="sxs-lookup"><span data-stu-id="c940b-390">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-391">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-392">1.3</span><span class="sxs-lookup"><span data-stu-id="c940b-392">1.3</span></span>|
|[<span data-ttu-id="c940b-393">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-393">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-394">ReadItem</span></span>|
|[<span data-ttu-id="c940b-395">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-395">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-396">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-396">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c940b-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c940b-398">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c940b-398">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c940b-399">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-399">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-400">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-400">Read mode</span></span>

<span data-ttu-id="c940b-401">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-401">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-402">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-402">Compose mode</span></span>

<span data-ttu-id="c940b-403">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-403">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-404">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-404">Type:</span></span>

*   <span data-ttu-id="c940b-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-406">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-406">Requirements</span></span>

|<span data-ttu-id="c940b-407">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-407">Requirement</span></span>| <span data-ttu-id="c940b-408">値</span><span class="sxs-lookup"><span data-stu-id="c940b-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-410">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-410">1.0</span></span>|
|[<span data-ttu-id="c940b-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-412">ReadItem</span></span>|
|[<span data-ttu-id="c940b-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-414">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-414">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-415">例</span><span class="sxs-lookup"><span data-stu-id="c940b-415">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c940b-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c940b-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c940b-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-419">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-419">Type:</span></span>

*   [<span data-ttu-id="c940b-420">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c940b-420">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c940b-421">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-421">Requirements</span></span>

|<span data-ttu-id="c940b-422">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-422">Requirement</span></span>| <span data-ttu-id="c940b-423">値</span><span class="sxs-lookup"><span data-stu-id="c940b-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-424">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-425">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-425">1.0</span></span>|
|[<span data-ttu-id="c940b-426">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-427">ReadItem</span></span>|
|[<span data-ttu-id="c940b-428">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-429">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-430">例</span><span class="sxs-lookup"><span data-stu-id="c940b-430">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c940b-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c940b-432">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c940b-432">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c940b-433">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-433">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-434">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-434">Read mode</span></span>

<span data-ttu-id="c940b-435">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-435">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-436">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-436">Compose mode</span></span>

<span data-ttu-id="c940b-437">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-437">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-438">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-438">Type:</span></span>

*   <span data-ttu-id="c940b-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-440">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-440">Requirements</span></span>

|<span data-ttu-id="c940b-441">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-441">Requirement</span></span>| <span data-ttu-id="c940b-442">値</span><span class="sxs-lookup"><span data-stu-id="c940b-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-443">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-444">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-444">1.0</span></span>|
|[<span data-ttu-id="c940b-445">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-446">ReadItem</span></span>|
|[<span data-ttu-id="c940b-447">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-448">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-449">例</span><span class="sxs-lookup"><span data-stu-id="c940b-449">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="c940b-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c940b-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="c940b-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="c940b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c940b-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-455">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="c940b-455">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-456">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-456">Type:</span></span>

*   [<span data-ttu-id="c940b-457">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c940b-457">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c940b-458">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-458">Requirements</span></span>

|<span data-ttu-id="c940b-459">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-459">Requirement</span></span>| <span data-ttu-id="c940b-460">値</span><span class="sxs-lookup"><span data-stu-id="c940b-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-461">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-462">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-462">1.0</span></span>|
|[<span data-ttu-id="c940b-463">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-464">ReadItem</span></span>|
|[<span data-ttu-id="c940b-465">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-466">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-466">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-467">例</span><span class="sxs-lookup"><span data-stu-id="c940b-467">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="c940b-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c940b-468">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="c940b-469">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-469">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c940b-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-472">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-472">Read mode</span></span>

<span data-ttu-id="c940b-473">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-473">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-474">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-474">Compose mode</span></span>

<span data-ttu-id="c940b-475">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-475">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c940b-476">[`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c940b-476">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-477">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-477">Type:</span></span>

*   <span data-ttu-id="c940b-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="c940b-478">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-479">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-479">Requirements</span></span>

|<span data-ttu-id="c940b-480">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-480">Requirement</span></span>| <span data-ttu-id="c940b-481">値</span><span class="sxs-lookup"><span data-stu-id="c940b-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-482">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-483">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-483">1.0</span></span>|
|[<span data-ttu-id="c940b-484">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-485">ReadItem</span></span>|
|[<span data-ttu-id="c940b-486">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-487">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-487">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-488">例</span><span class="sxs-lookup"><span data-stu-id="c940b-488">Example</span></span>

<span data-ttu-id="c940b-489">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-489">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="c940b-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c940b-490">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="c940b-491">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c940b-492">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="c940b-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-493">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-493">Read mode</span></span>

<span data-ttu-id="c940b-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c940b-496">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-496">Compose mode</span></span>

<span data-ttu-id="c940b-497">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c940b-498">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-498">Type:</span></span>

*   <span data-ttu-id="c940b-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c940b-499">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-500">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-500">Requirements</span></span>

|<span data-ttu-id="c940b-501">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-501">Requirement</span></span>| <span data-ttu-id="c940b-502">値</span><span class="sxs-lookup"><span data-stu-id="c940b-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-503">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-504">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-504">1.0</span></span>|
|[<span data-ttu-id="c940b-505">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-506">ReadItem</span></span>|
|[<span data-ttu-id="c940b-507">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-508">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-508">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="c940b-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="c940b-510">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c940b-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c940b-511">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c940b-512">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="c940b-512">Read mode</span></span>

<span data-ttu-id="c940b-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c940b-515">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="c940b-515">Compose mode</span></span>

<span data-ttu-id="c940b-516">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c940b-517">型:</span><span class="sxs-lookup"><span data-stu-id="c940b-517">Type:</span></span>

*   <span data-ttu-id="c940b-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c940b-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-519">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-519">Requirements</span></span>

|<span data-ttu-id="c940b-520">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-520">Requirement</span></span>| <span data-ttu-id="c940b-521">値</span><span class="sxs-lookup"><span data-stu-id="c940b-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-522">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-523">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-523">1.0</span></span>|
|[<span data-ttu-id="c940b-524">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-525">ReadItem</span></span>|
|[<span data-ttu-id="c940b-526">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-527">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-528">例</span><span class="sxs-lookup"><span data-stu-id="c940b-528">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c940b-529">メソッド</span><span class="sxs-lookup"><span data-stu-id="c940b-529">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c940b-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c940b-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c940b-531">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c940b-531">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c940b-532">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="c940b-532">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c940b-533">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-533">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-534">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-534">Parameters:</span></span>

|<span data-ttu-id="c940b-535">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-535">Name</span></span>| <span data-ttu-id="c940b-536">型</span><span class="sxs-lookup"><span data-stu-id="c940b-536">Type</span></span>| <span data-ttu-id="c940b-537">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-537">Attributes</span></span>| <span data-ttu-id="c940b-538">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-538">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c940b-539">String</span><span class="sxs-lookup"><span data-stu-id="c940b-539">String</span></span>||<span data-ttu-id="c940b-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c940b-542">String</span><span class="sxs-lookup"><span data-stu-id="c940b-542">String</span></span>||<span data-ttu-id="c940b-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c940b-545">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-545">Object</span></span>| <span data-ttu-id="c940b-546">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-546">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-547">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-547">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-548">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-548">Object</span></span>| <span data-ttu-id="c940b-549">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-549">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-550">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-550">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c940b-551">function</span><span class="sxs-lookup"><span data-stu-id="c940b-551">function</span></span>| <span data-ttu-id="c940b-552">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-552">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-553">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c940b-554">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-554">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c940b-555">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-555">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c940b-556">エラー</span><span class="sxs-lookup"><span data-stu-id="c940b-556">Errors</span></span>

| <span data-ttu-id="c940b-557">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c940b-557">Error code</span></span> | <span data-ttu-id="c940b-558">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-558">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c940b-559">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="c940b-559">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c940b-560">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="c940b-560">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c940b-561">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c940b-561">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-562">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-562">Requirements</span></span>

|<span data-ttu-id="c940b-563">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-563">Requirement</span></span>| <span data-ttu-id="c940b-564">値</span><span class="sxs-lookup"><span data-stu-id="c940b-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-565">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-566">1.1</span><span class="sxs-lookup"><span data-stu-id="c940b-566">1.1</span></span>|
|[<span data-ttu-id="c940b-567">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-568">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-568">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-569">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-570">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-570">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-571">例</span><span class="sxs-lookup"><span data-stu-id="c940b-571">Example</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c940b-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c940b-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c940b-573">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="c940b-573">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c940b-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c940b-577">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-577">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c940b-578">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="c940b-578">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-579">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-579">Parameters:</span></span>

|<span data-ttu-id="c940b-580">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-580">Name</span></span>| <span data-ttu-id="c940b-581">型</span><span class="sxs-lookup"><span data-stu-id="c940b-581">Type</span></span>| <span data-ttu-id="c940b-582">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-582">Attributes</span></span>| <span data-ttu-id="c940b-583">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-583">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c940b-584">String</span><span class="sxs-lookup"><span data-stu-id="c940b-584">String</span></span>||<span data-ttu-id="c940b-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c940b-587">String</span><span class="sxs-lookup"><span data-stu-id="c940b-587">String</span></span>||<span data-ttu-id="c940b-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c940b-590">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-590">Object</span></span>| <span data-ttu-id="c940b-591">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-591">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-592">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-592">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-593">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-593">Object</span></span>| <span data-ttu-id="c940b-594">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-594">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-595">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-595">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c940b-596">function</span><span class="sxs-lookup"><span data-stu-id="c940b-596">function</span></span>| <span data-ttu-id="c940b-597">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-597">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-598">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c940b-599">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-599">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c940b-600">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-600">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c940b-601">エラー</span><span class="sxs-lookup"><span data-stu-id="c940b-601">Errors</span></span>

| <span data-ttu-id="c940b-602">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c940b-602">Error code</span></span> | <span data-ttu-id="c940b-603">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-603">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c940b-604">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="c940b-604">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-605">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-605">Requirements</span></span>

|<span data-ttu-id="c940b-606">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-606">Requirement</span></span>| <span data-ttu-id="c940b-607">値</span><span class="sxs-lookup"><span data-stu-id="c940b-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-608">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-609">1.1</span><span class="sxs-lookup"><span data-stu-id="c940b-609">1.1</span></span>|
|[<span data-ttu-id="c940b-610">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-612">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-613">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-613">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-614">例</span><span class="sxs-lookup"><span data-stu-id="c940b-614">Example</span></span>

<span data-ttu-id="c940b-615">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-615">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="c940b-616">close()</span><span class="sxs-lookup"><span data-stu-id="c940b-616">close()</span></span>

<span data-ttu-id="c940b-617">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="c940b-617">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c940b-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-620">Outlook on the web で、予定のアイテムが `saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、キャンセルのいずれかを行うようダイアログが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-620">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c940b-621">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="c940b-621">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-622">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-622">Requirements</span></span>

|<span data-ttu-id="c940b-623">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-623">Requirement</span></span>| <span data-ttu-id="c940b-624">値</span><span class="sxs-lookup"><span data-stu-id="c940b-624">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-625">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-625">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-626">1.3</span><span class="sxs-lookup"><span data-stu-id="c940b-626">1.3</span></span>|
|[<span data-ttu-id="c940b-627">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-627">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-628">制限あり</span><span class="sxs-lookup"><span data-stu-id="c940b-628">Restricted</span></span>|
|[<span data-ttu-id="c940b-629">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-629">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-630">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-630">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c940b-631">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c940b-631">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c940b-632">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-632">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-633">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-633">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c940b-634">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-634">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c940b-635">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c940b-635">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c940b-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c940b-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-639">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-639">Parameters:</span></span>

|<span data-ttu-id="c940b-640">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-640">Name</span></span>| <span data-ttu-id="c940b-641">型</span><span class="sxs-lookup"><span data-stu-id="c940b-641">Type</span></span>| <span data-ttu-id="c940b-642">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-642">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c940b-643">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c940b-643">String &#124; Object</span></span>| |<span data-ttu-id="c940b-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c940b-646">**または**</span><span class="sxs-lookup"><span data-stu-id="c940b-646">**OR**</span></span><br/><span data-ttu-id="c940b-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c940b-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c940b-649">String</span><span class="sxs-lookup"><span data-stu-id="c940b-649">String</span></span> | <span data-ttu-id="c940b-650">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-650">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c940b-653">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-653">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c940b-654">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-654">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-655">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c940b-655">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c940b-656">String</span><span class="sxs-lookup"><span data-stu-id="c940b-656">String</span></span> | | <span data-ttu-id="c940b-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c940b-659">String</span><span class="sxs-lookup"><span data-stu-id="c940b-659">String</span></span> | | <span data-ttu-id="c940b-660">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-660">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c940b-661">String</span><span class="sxs-lookup"><span data-stu-id="c940b-661">String</span></span> | | <span data-ttu-id="c940b-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c940b-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c940b-664">String</span><span class="sxs-lookup"><span data-stu-id="c940b-664">String</span></span> | | <span data-ttu-id="c940b-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c940b-668">function</span><span class="sxs-lookup"><span data-stu-id="c940b-668">function</span></span> | <span data-ttu-id="c940b-669">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-669">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-670">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-671">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-671">Requirements</span></span>

|<span data-ttu-id="c940b-672">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-672">Requirement</span></span>| <span data-ttu-id="c940b-673">値</span><span class="sxs-lookup"><span data-stu-id="c940b-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-674">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-675">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-675">1.0</span></span>|
|[<span data-ttu-id="c940b-676">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-676">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-677">ReadItem</span></span>|
|[<span data-ttu-id="c940b-678">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-678">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-679">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-679">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c940b-680">例</span><span class="sxs-lookup"><span data-stu-id="c940b-680">Examples</span></span>

<span data-ttu-id="c940b-681">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c940b-681">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c940b-682">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-682">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c940b-683">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-683">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c940b-684">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-684">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c940b-685">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-685">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c940b-686">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-686">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c940b-687">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c940b-687">displayReplyForm(formData)</span></span>

<span data-ttu-id="c940b-688">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-688">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-689">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-689">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c940b-690">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-690">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c940b-691">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="c940b-691">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c940b-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="c940b-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-695">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-695">Parameters:</span></span>

|<span data-ttu-id="c940b-696">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-696">Name</span></span>| <span data-ttu-id="c940b-697">型</span><span class="sxs-lookup"><span data-stu-id="c940b-697">Type</span></span>| <span data-ttu-id="c940b-698">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-698">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c940b-699">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c940b-699">String &#124; Object</span></span>| | <span data-ttu-id="c940b-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c940b-702">**または**</span><span class="sxs-lookup"><span data-stu-id="c940b-702">**OR**</span></span><br/><span data-ttu-id="c940b-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="c940b-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c940b-705">String</span><span class="sxs-lookup"><span data-stu-id="c940b-705">String</span></span> | <span data-ttu-id="c940b-706">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-706">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="c940b-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c940b-709">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-709">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c940b-710">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-710">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-711">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="c940b-711">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c940b-712">String</span><span class="sxs-lookup"><span data-stu-id="c940b-712">String</span></span> | | <span data-ttu-id="c940b-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c940b-715">String</span><span class="sxs-lookup"><span data-stu-id="c940b-715">String</span></span> | | <span data-ttu-id="c940b-716">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="c940b-716">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c940b-717">String</span><span class="sxs-lookup"><span data-stu-id="c940b-717">String</span></span> | | <span data-ttu-id="c940b-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="c940b-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c940b-720">String</span><span class="sxs-lookup"><span data-stu-id="c940b-720">String</span></span> | | <span data-ttu-id="c940b-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="c940b-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c940b-724">function</span><span class="sxs-lookup"><span data-stu-id="c940b-724">function</span></span> | <span data-ttu-id="c940b-725">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-725">&lt;optional&gt;</span></span> | <span data-ttu-id="c940b-726">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-726">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-727">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-727">Requirements</span></span>

|<span data-ttu-id="c940b-728">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-728">Requirement</span></span>| <span data-ttu-id="c940b-729">値</span><span class="sxs-lookup"><span data-stu-id="c940b-729">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-730">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-730">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-731">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-731">1.0</span></span>|
|[<span data-ttu-id="c940b-732">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-732">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-733">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-733">ReadItem</span></span>|
|[<span data-ttu-id="c940b-734">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-734">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-735">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-735">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c940b-736">例</span><span class="sxs-lookup"><span data-stu-id="c940b-736">Examples</span></span>

<span data-ttu-id="c940b-737">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="c940b-737">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c940b-738">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-738">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c940b-739">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-739">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c940b-740">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-740">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c940b-741">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-741">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c940b-742">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="c940b-742">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="c940b-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c940b-743">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="c940b-744">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-744">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-745">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-745">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-746">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-746">Requirements</span></span>

|<span data-ttu-id="c940b-747">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-747">Requirement</span></span>| <span data-ttu-id="c940b-748">値</span><span class="sxs-lookup"><span data-stu-id="c940b-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-749">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-750">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-750">1.0</span></span>|
|[<span data-ttu-id="c940b-751">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-752">ReadItem</span></span>|
|[<span data-ttu-id="c940b-753">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-754">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-755">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-755">Returns:</span></span>

<span data-ttu-id="c940b-756">型:[Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c940b-756">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c940b-757">例</span><span class="sxs-lookup"><span data-stu-id="c940b-757">Example</span></span>

<span data-ttu-id="c940b-758">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="c940b-758">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="c940b-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c940b-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c940b-760">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-760">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-761">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-762">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-762">Parameters:</span></span>

|<span data-ttu-id="c940b-763">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-763">Name</span></span>| <span data-ttu-id="c940b-764">型</span><span class="sxs-lookup"><span data-stu-id="c940b-764">Type</span></span>| <span data-ttu-id="c940b-765">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-765">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c940b-766">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c940b-766">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="c940b-767">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="c940b-767">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-768">Requirements</span><span class="sxs-lookup"><span data-stu-id="c940b-768">Requirements</span></span>

|<span data-ttu-id="c940b-769">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-769">Requirement</span></span>| <span data-ttu-id="c940b-770">値</span><span class="sxs-lookup"><span data-stu-id="c940b-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-771">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-772">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-772">1.0</span></span>|
|[<span data-ttu-id="c940b-773">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-774">制限あり</span><span class="sxs-lookup"><span data-stu-id="c940b-774">Restricted</span></span>|
|[<span data-ttu-id="c940b-775">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-776">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-777">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-777">Returns:</span></span>

<span data-ttu-id="c940b-778">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-778">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c940b-779">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-779">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c940b-780">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-780">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c940b-781">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="c940b-781">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c940b-782">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="c940b-782">Value of `entityType`</span></span> | <span data-ttu-id="c940b-783">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="c940b-783">Type of objects in returned array</span></span> | <span data-ttu-id="c940b-784">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="c940b-784">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c940b-785">文字列</span><span class="sxs-lookup"><span data-stu-id="c940b-785">String</span></span> | <span data-ttu-id="c940b-786">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c940b-786">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c940b-787">連絡先</span><span class="sxs-lookup"><span data-stu-id="c940b-787">Contact</span></span> | <span data-ttu-id="c940b-788">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c940b-788">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c940b-789">文字列</span><span class="sxs-lookup"><span data-stu-id="c940b-789">String</span></span> | <span data-ttu-id="c940b-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c940b-790">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c940b-791">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c940b-791">MeetingSuggestion</span></span> | <span data-ttu-id="c940b-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c940b-792">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c940b-793">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c940b-793">PhoneNumber</span></span> | <span data-ttu-id="c940b-794">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c940b-794">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c940b-795">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c940b-795">TaskSuggestion</span></span> | <span data-ttu-id="c940b-796">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c940b-796">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c940b-797">文字列</span><span class="sxs-lookup"><span data-stu-id="c940b-797">String</span></span> | <span data-ttu-id="c940b-798">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="c940b-798">**Restricted**</span></span> |

<span data-ttu-id="c940b-799">型:Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c940b-799">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c940b-800">例</span><span class="sxs-lookup"><span data-stu-id="c940b-800">Example</span></span>

<span data-ttu-id="c940b-801">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="c940b-801">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="c940b-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c940b-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c940b-803">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-803">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-804">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-804">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c940b-805">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-805">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-806">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-806">Parameters:</span></span>

|<span data-ttu-id="c940b-807">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-807">Name</span></span>| <span data-ttu-id="c940b-808">型</span><span class="sxs-lookup"><span data-stu-id="c940b-808">Type</span></span>| <span data-ttu-id="c940b-809">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-809">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c940b-810">String</span><span class="sxs-lookup"><span data-stu-id="c940b-810">String</span></span>|<span data-ttu-id="c940b-811">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c940b-811">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-812">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-812">Requirements</span></span>

|<span data-ttu-id="c940b-813">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-813">Requirement</span></span>| <span data-ttu-id="c940b-814">値</span><span class="sxs-lookup"><span data-stu-id="c940b-814">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-815">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-815">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-816">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-816">1.0</span></span>|
|[<span data-ttu-id="c940b-817">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-817">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-818">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-818">ReadItem</span></span>|
|[<span data-ttu-id="c940b-819">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-819">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-820">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-820">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-821">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-821">Returns:</span></span>

<span data-ttu-id="c940b-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c940b-824">型:Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c940b-824">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c940b-825">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c940b-825">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c940b-826">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-826">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-827">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-827">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c940b-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c940b-831">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="c940b-831">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c940b-832">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="c940b-832">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c940b-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c940b-836">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-836">Requirements</span></span>

|<span data-ttu-id="c940b-837">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-837">Requirement</span></span>| <span data-ttu-id="c940b-838">値</span><span class="sxs-lookup"><span data-stu-id="c940b-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-839">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-840">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-840">1.0</span></span>|
|[<span data-ttu-id="c940b-841">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-842">ReadItem</span></span>|
|[<span data-ttu-id="c940b-843">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-844">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-845">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-845">Returns:</span></span>

<span data-ttu-id="c940b-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="c940b-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c940b-848">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c940b-848">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c940b-849">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-849">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c940b-850">例</span><span class="sxs-lookup"><span data-stu-id="c940b-850">Example</span></span>

<span data-ttu-id="c940b-851">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="c940b-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c940b-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c940b-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c940b-853">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-854">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-854">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c940b-855">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="c940b-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c940b-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="c940b-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-858">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-858">Parameters:</span></span>

|<span data-ttu-id="c940b-859">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-859">Name</span></span>| <span data-ttu-id="c940b-860">型</span><span class="sxs-lookup"><span data-stu-id="c940b-860">Type</span></span>| <span data-ttu-id="c940b-861">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c940b-862">String</span><span class="sxs-lookup"><span data-stu-id="c940b-862">String</span></span>|<span data-ttu-id="c940b-863">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="c940b-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-864">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-864">Requirements</span></span>

|<span data-ttu-id="c940b-865">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-865">Requirement</span></span>| <span data-ttu-id="c940b-866">値</span><span class="sxs-lookup"><span data-stu-id="c940b-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-867">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-868">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-868">1.0</span></span>|
|[<span data-ttu-id="c940b-869">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-869">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-870">ReadItem</span></span>|
|[<span data-ttu-id="c940b-871">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-871">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-872">読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-873">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-873">Returns:</span></span>

<span data-ttu-id="c940b-874">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="c940b-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c940b-875">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c940b-875">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c940b-876">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c940b-876">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c940b-877">例</span><span class="sxs-lookup"><span data-stu-id="c940b-877">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c940b-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c940b-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c940b-879">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-879">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c940b-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-882">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-882">Parameters:</span></span>

|<span data-ttu-id="c940b-883">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-883">Name</span></span>| <span data-ttu-id="c940b-884">型</span><span class="sxs-lookup"><span data-stu-id="c940b-884">Type</span></span>| <span data-ttu-id="c940b-885">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-885">Attributes</span></span>| <span data-ttu-id="c940b-886">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-886">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c940b-887">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c940b-887">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c940b-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c940b-891">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-891">Object</span></span>| <span data-ttu-id="c940b-892">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-892">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-893">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-893">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-894">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-894">Object</span></span>| <span data-ttu-id="c940b-895">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-895">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-896">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-896">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c940b-897">function</span><span class="sxs-lookup"><span data-stu-id="c940b-897">function</span></span>||<span data-ttu-id="c940b-898">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-898">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c940b-899">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="c940b-899">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c940b-900">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="c940b-900">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-901">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-901">Requirements</span></span>

|<span data-ttu-id="c940b-902">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-902">Requirement</span></span>| <span data-ttu-id="c940b-903">値</span><span class="sxs-lookup"><span data-stu-id="c940b-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-904">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-905">1.2</span><span class="sxs-lookup"><span data-stu-id="c940b-905">1.2</span></span>|
|[<span data-ttu-id="c940b-906">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-907">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-907">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-908">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-909">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-909">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c940b-910">戻り値:</span><span class="sxs-lookup"><span data-stu-id="c940b-910">Returns:</span></span>

<span data-ttu-id="c940b-911">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="c940b-911">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c940b-912">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c940b-912">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c940b-913">String</span><span class="sxs-lookup"><span data-stu-id="c940b-913">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c940b-914">例</span><span class="sxs-lookup"><span data-stu-id="c940b-914">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c940b-915">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c940b-915">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c940b-916">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="c940b-916">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c940b-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="c940b-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-920">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-920">Parameters:</span></span>

|<span data-ttu-id="c940b-921">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-921">Name</span></span>| <span data-ttu-id="c940b-922">型</span><span class="sxs-lookup"><span data-stu-id="c940b-922">Type</span></span>| <span data-ttu-id="c940b-923">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-923">Attributes</span></span>| <span data-ttu-id="c940b-924">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-924">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c940b-925">function</span><span class="sxs-lookup"><span data-stu-id="c940b-925">function</span></span>||<span data-ttu-id="c940b-926">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-926">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c940b-927">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-927">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c940b-928">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-928">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c940b-929">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c940b-929">Object</span></span>| <span data-ttu-id="c940b-930">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-930">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-931">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-931">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c940b-932">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="c940b-932">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-933">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-933">Requirements</span></span>

|<span data-ttu-id="c940b-934">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-934">Requirement</span></span>| <span data-ttu-id="c940b-935">値</span><span class="sxs-lookup"><span data-stu-id="c940b-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-936">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-937">1.0</span><span class="sxs-lookup"><span data-stu-id="c940b-937">1.0</span></span>|
|[<span data-ttu-id="c940b-938">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c940b-939">ReadItem</span></span>|
|[<span data-ttu-id="c940b-940">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-941">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="c940b-941">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-942">例</span><span class="sxs-lookup"><span data-stu-id="c940b-942">Example</span></span>

<span data-ttu-id="c940b-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c940b-946">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c940b-946">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c940b-947">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c940b-947">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c940b-p165">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="c940b-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-952">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-952">Parameters:</span></span>

|<span data-ttu-id="c940b-953">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-953">Name</span></span>| <span data-ttu-id="c940b-954">型</span><span class="sxs-lookup"><span data-stu-id="c940b-954">Type</span></span>| <span data-ttu-id="c940b-955">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-955">Attributes</span></span>| <span data-ttu-id="c940b-956">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-956">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c940b-957">String</span><span class="sxs-lookup"><span data-stu-id="c940b-957">String</span></span>||<span data-ttu-id="c940b-958">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="c940b-958">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="c940b-959">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c940b-959">Object</span></span>| <span data-ttu-id="c940b-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-960">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-961">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-962">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-962">Object</span></span>| <span data-ttu-id="c940b-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-963">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-964">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c940b-965">function</span><span class="sxs-lookup"><span data-stu-id="c940b-965">function</span></span>| <span data-ttu-id="c940b-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-966">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-967">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c940b-968">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c940b-969">エラー</span><span class="sxs-lookup"><span data-stu-id="c940b-969">Errors</span></span>

| <span data-ttu-id="c940b-970">エラー コード</span><span class="sxs-lookup"><span data-stu-id="c940b-970">Error code</span></span> | <span data-ttu-id="c940b-971">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c940b-972">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="c940b-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-973">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-973">Requirements</span></span>

|<span data-ttu-id="c940b-974">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-974">Requirement</span></span>| <span data-ttu-id="c940b-975">値</span><span class="sxs-lookup"><span data-stu-id="c940b-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-976">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-977">1.1</span><span class="sxs-lookup"><span data-stu-id="c940b-977">1.1</span></span>|
|[<span data-ttu-id="c940b-978">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-980">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-981">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-982">例</span><span class="sxs-lookup"><span data-stu-id="c940b-982">Example</span></span>

<span data-ttu-id="c940b-983">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="c940b-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c940b-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c940b-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="c940b-985">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="c940b-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="c940b-p166">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-989">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="c940b-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c940b-990">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c940b-p168">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c940b-994">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="c940b-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c940b-995">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="c940b-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c940b-996">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c940b-997">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-998">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-998">Parameters:</span></span>

|<span data-ttu-id="c940b-999">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-999">Name</span></span>| <span data-ttu-id="c940b-1000">型</span><span class="sxs-lookup"><span data-stu-id="c940b-1000">Type</span></span>| <span data-ttu-id="c940b-1001">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-1001">Attributes</span></span>| <span data-ttu-id="c940b-1002">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="c940b-1003">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c940b-1003">Object</span></span>| <span data-ttu-id="c940b-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-1005">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-1006">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-1006">Object</span></span>| <span data-ttu-id="c940b-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-1008">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c940b-1009">function</span><span class="sxs-lookup"><span data-stu-id="c940b-1009">function</span></span>||<span data-ttu-id="c940b-1010">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c940b-1011">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c940b-1012">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-1012">Requirements</span></span>

|<span data-ttu-id="c940b-1013">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-1013">Requirement</span></span>| <span data-ttu-id="c940b-1014">値</span><span class="sxs-lookup"><span data-stu-id="c940b-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-1015">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="c940b-1016">1.3</span></span>|
|[<span data-ttu-id="c940b-1017">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-1019">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-1020">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c940b-1021">例</span><span class="sxs-lookup"><span data-stu-id="c940b-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c940b-p170">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c940b-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c940b-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c940b-1025">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="c940b-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c940b-p171">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c940b-1029">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="c940b-1029">Parameters:</span></span>

|<span data-ttu-id="c940b-1030">名前</span><span class="sxs-lookup"><span data-stu-id="c940b-1030">Name</span></span>| <span data-ttu-id="c940b-1031">型</span><span class="sxs-lookup"><span data-stu-id="c940b-1031">Type</span></span>| <span data-ttu-id="c940b-1032">属性</span><span class="sxs-lookup"><span data-stu-id="c940b-1032">Attributes</span></span>| <span data-ttu-id="c940b-1033">説明</span><span class="sxs-lookup"><span data-stu-id="c940b-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c940b-1034">String</span><span class="sxs-lookup"><span data-stu-id="c940b-1034">String</span></span>||<span data-ttu-id="c940b-p172">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c940b-1038">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-1038">Object</span></span>| <span data-ttu-id="c940b-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-1040">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="c940b-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c940b-1041">Object</span><span class="sxs-lookup"><span data-stu-id="c940b-1041">Object</span></span>| <span data-ttu-id="c940b-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-1043">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="c940b-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c940b-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="c940b-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c940b-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="c940b-p173">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c940b-p174">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c940b-1050">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c940b-1051">function</span><span class="sxs-lookup"><span data-stu-id="c940b-1051">function</span></span>||<span data-ttu-id="c940b-1052">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="c940b-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c940b-1053">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-1053">Requirements</span></span>

|<span data-ttu-id="c940b-1054">要件</span><span class="sxs-lookup"><span data-stu-id="c940b-1054">Requirement</span></span>| <span data-ttu-id="c940b-1055">値</span><span class="sxs-lookup"><span data-stu-id="c940b-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="c940b-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c940b-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c940b-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="c940b-1057">1.2</span></span>|
|[<span data-ttu-id="c940b-1058">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c940b-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c940b-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c940b-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="c940b-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c940b-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c940b-1061">作成</span><span class="sxs-lookup"><span data-stu-id="c940b-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c940b-1062">例</span><span class="sxs-lookup"><span data-stu-id="c940b-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
