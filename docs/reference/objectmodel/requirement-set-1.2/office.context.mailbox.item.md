---
title: Office.context.mailbox.item ・要件設定 1.2
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 2ac3df2a8daae00e64bb66247e66834f9da4243c
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701870"
---
# <a name="item"></a><span data-ttu-id="a6c49-102">item</span><span class="sxs-lookup"><span data-stu-id="a6c49-102">item</span></span>

### <span data-ttu-id="a6c49-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="a6c49-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="a6c49-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-107">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-107">Requirements</span></span>

|<span data-ttu-id="a6c49-108">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-108">Requirement</span></span>| <span data-ttu-id="a6c49-109">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-111">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-111">1.0</span></span>|
|[<span data-ttu-id="a6c49-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="a6c49-113">Restricted</span></span>|
|[<span data-ttu-id="a6c49-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-115">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="a6c49-116">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-116">Example</span></span>

<span data-ttu-id="a6c49-117">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="a6c49-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="a6c49-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="a6c49-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6c49-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="a6c49-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-122">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="a6c49-123">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6c49-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-124">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-124">Type:</span></span>

*   <span data-ttu-id="a6c49-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="a6c49-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-126">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-126">Requirements</span></span>

|<span data-ttu-id="a6c49-127">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-127">Requirement</span></span>| <span data-ttu-id="a6c49-128">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-129">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-130">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-130">1.0</span></span>|
|[<span data-ttu-id="a6c49-131">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-132">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-134">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-135">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-135">Example</span></span>

<span data-ttu-id="a6c49-136">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="a6c49-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="a6c49-138">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="a6c49-139">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-140">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-140">Type:</span></span>

*   [<span data-ttu-id="a6c49-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="a6c49-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="a6c49-142">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-142">Requirements</span></span>

|<span data-ttu-id="a6c49-143">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-143">Requirement</span></span>| <span data-ttu-id="a6c49-144">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-146">1.1</span><span class="sxs-lookup"><span data-stu-id="a6c49-146">1.1</span></span>|
|[<span data-ttu-id="a6c49-147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-148">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-150">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-151">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="a6c49-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="a6c49-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="a6c49-153">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-154">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-154">Type:</span></span>

*   [<span data-ttu-id="a6c49-155">Body</span><span class="sxs-lookup"><span data-stu-id="a6c49-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="a6c49-156">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-156">Requirements</span></span>

|<span data-ttu-id="a6c49-157">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-157">Requirement</span></span>| <span data-ttu-id="a6c49-158">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-160">1.1</span><span class="sxs-lookup"><span data-stu-id="a6c49-160">1.1</span></span>|
|[<span data-ttu-id="a6c49-161">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-162">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-164">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="a6c49-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="a6c49-166">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="a6c49-167">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-168">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-168">Read mode</span></span>

<span data-ttu-id="a6c49-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-171">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-171">Compose mode</span></span>

<span data-ttu-id="a6c49-172">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-173">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-173">Type:</span></span>

*   <span data-ttu-id="a6c49-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-175">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-175">Requirements</span></span>

|<span data-ttu-id="a6c49-176">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-176">Requirement</span></span>| <span data-ttu-id="a6c49-177">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-179">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-179">1.0</span></span>|
|[<span data-ttu-id="a6c49-180">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-181">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-183">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-184">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="a6c49-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="a6c49-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="a6c49-186">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="a6c49-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="a6c49-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-191">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-191">Type:</span></span>

*   <span data-ttu-id="a6c49-192">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-193">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-193">Requirements</span></span>

|<span data-ttu-id="a6c49-194">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-194">Requirement</span></span>| <span data-ttu-id="a6c49-195">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-196">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-197">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-197">1.0</span></span>|
|[<span data-ttu-id="a6c49-198">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-199">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="a6c49-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="a6c49-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="a6c49-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="a6c49-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-205">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-205">Type:</span></span>

*   <span data-ttu-id="a6c49-206">日付</span><span class="sxs-lookup"><span data-stu-id="a6c49-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-207">Requirements</span></span>

|<span data-ttu-id="a6c49-208">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-208">Requirement</span></span>| <span data-ttu-id="a6c49-209">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-211">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-211">1.0</span></span>|
|[<span data-ttu-id="a6c49-212">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-213">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-216">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="a6c49-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="a6c49-217">dateTimeModified :Date</span></span>

<span data-ttu-id="a6c49-p111">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-220">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-221">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-221">Type:</span></span>

*   <span data-ttu-id="a6c49-222">日付</span><span class="sxs-lookup"><span data-stu-id="a6c49-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-223">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-223">Requirements</span></span>

|<span data-ttu-id="a6c49-224">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-224">Requirement</span></span>| <span data-ttu-id="a6c49-225">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-227">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-227">1.0</span></span>|
|[<span data-ttu-id="a6c49-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-229">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-231">Read</span><span class="sxs-lookup"><span data-stu-id="a6c49-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-232">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="a6c49-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6c49-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="a6c49-234">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="a6c49-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-237">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-237">Read mode</span></span>

<span data-ttu-id="a6c49-238">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-239">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-239">Compose mode</span></span>

<span data-ttu-id="a6c49-240">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="a6c49-241">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-242">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-242">Type:</span></span>

*   <span data-ttu-id="a6c49-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6c49-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-244">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-244">Requirements</span></span>

|<span data-ttu-id="a6c49-245">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-245">Requirement</span></span>| <span data-ttu-id="a6c49-246">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-247">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-248">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-248">1.0</span></span>|
|[<span data-ttu-id="a6c49-249">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-250">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-251">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-252">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-253">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-253">Example</span></span>

<span data-ttu-id="a6c49-254">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="a6c49-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a6c49-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="a6c49-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="a6c49-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-260">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-261">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-261">Type:</span></span>

*   [<span data-ttu-id="a6c49-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a6c49-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a6c49-263">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-263">Requirements</span></span>

|<span data-ttu-id="a6c49-264">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-264">Requirement</span></span>| <span data-ttu-id="a6c49-265">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-267">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-267">1.0</span></span>|
|[<span data-ttu-id="a6c49-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-269">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-271">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="a6c49-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="a6c49-272">internetMessageId :String</span></span>

<span data-ttu-id="a6c49-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-275">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-275">Type:</span></span>

*   <span data-ttu-id="a6c49-276">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-277">Requirements</span></span>

|<span data-ttu-id="a6c49-278">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-278">Requirement</span></span>| <span data-ttu-id="a6c49-279">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-281">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-281">1.0</span></span>|
|[<span data-ttu-id="a6c49-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-283">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-285">Read</span><span class="sxs-lookup"><span data-stu-id="a6c49-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-286">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="a6c49-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="a6c49-287">itemClass :String</span></span>

<span data-ttu-id="a6c49-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="a6c49-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="a6c49-292">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-292">Type</span></span> | <span data-ttu-id="a6c49-293">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-293">Description</span></span> | <span data-ttu-id="a6c49-294">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="a6c49-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="a6c49-295">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="a6c49-295">Appointment items</span></span> | <span data-ttu-id="a6c49-296">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="a6c49-297">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="a6c49-297">Message items</span></span> | <span data-ttu-id="a6c49-298">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="a6c49-299">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-300">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-300">Type:</span></span>

*   <span data-ttu-id="a6c49-301">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-302">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-302">Requirements</span></span>

|<span data-ttu-id="a6c49-303">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-303">Requirement</span></span>| <span data-ttu-id="a6c49-304">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-305">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-306">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-306">1.0</span></span>|
|[<span data-ttu-id="a6c49-307">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-308">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-310">Read</span><span class="sxs-lookup"><span data-stu-id="a6c49-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-311">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="a6c49-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="a6c49-312">(nullable) itemId :String</span></span>

<span data-ttu-id="a6c49-p118">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-315">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="a6c49-316">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="a6c49-317">この値を使用して REST API を呼び出す前に、要件セット 1.3 から使用できる `Office.context.mailbox.convertToRestId` を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="a6c49-318">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a6c49-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-319">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-319">Type:</span></span>

*   <span data-ttu-id="a6c49-320">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-321">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-321">Requirements</span></span>

|<span data-ttu-id="a6c49-322">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-322">Requirement</span></span>| <span data-ttu-id="a6c49-323">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-325">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-325">1.0</span></span>|
|[<span data-ttu-id="a6c49-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-327">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-329">Read</span><span class="sxs-lookup"><span data-stu-id="a6c49-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-330">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-330">Example</span></span>

<span data-ttu-id="a6c49-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="a6c49-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="a6c49-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="a6c49-334">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="a6c49-335">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-336">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-336">Type:</span></span>

*   [<span data-ttu-id="a6c49-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="a6c49-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="a6c49-338">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-338">Requirements</span></span>

|<span data-ttu-id="a6c49-339">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-339">Requirement</span></span>| <span data-ttu-id="a6c49-340">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-342">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-342">1.0</span></span>|
|[<span data-ttu-id="a6c49-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-344">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-346">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-347">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="a6c49-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="a6c49-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="a6c49-349">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-350">Read mode</span></span>

<span data-ttu-id="a6c49-351">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-352">Compose mode</span></span>

<span data-ttu-id="a6c49-353">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-354">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-354">Type:</span></span>

*   <span data-ttu-id="a6c49-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="a6c49-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-356">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-356">Requirements</span></span>

|<span data-ttu-id="a6c49-357">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-357">Requirement</span></span>| <span data-ttu-id="a6c49-358">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-360">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-360">1.0</span></span>|
|[<span data-ttu-id="a6c49-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-362">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-364">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-365">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="a6c49-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="a6c49-366">normalizedSubject :String</span></span>

<span data-ttu-id="a6c49-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="a6c49-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-371">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-371">Type:</span></span>

*   <span data-ttu-id="a6c49-372">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-373">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-373">Requirements</span></span>

|<span data-ttu-id="a6c49-374">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-374">Requirement</span></span>| <span data-ttu-id="a6c49-375">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-376">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-377">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-377">1.0</span></span>|
|[<span data-ttu-id="a6c49-378">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-379">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-381">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-382">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="a6c49-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="a6c49-384">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="a6c49-385">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-386">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-386">Read mode</span></span>

<span data-ttu-id="a6c49-387">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-388">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-388">Compose mode</span></span>

<span data-ttu-id="a6c49-389">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-390">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-390">Type:</span></span>

*   <span data-ttu-id="a6c49-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-392">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-392">Requirements</span></span>

|<span data-ttu-id="a6c49-393">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-393">Requirement</span></span>| <span data-ttu-id="a6c49-394">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-395">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-396">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-396">1.0</span></span>|
|[<span data-ttu-id="a6c49-397">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-398">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-400">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-401">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="a6c49-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a6c49-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="a6c49-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-405">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-405">Type:</span></span>

*   [<span data-ttu-id="a6c49-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a6c49-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a6c49-407">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-407">Requirements</span></span>

|<span data-ttu-id="a6c49-408">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-408">Requirement</span></span>| <span data-ttu-id="a6c49-409">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-411">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-411">1.0</span></span>|
|[<span data-ttu-id="a6c49-412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-413">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-415">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-416">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="a6c49-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="a6c49-418">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="a6c49-419">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-420">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-420">Read mode</span></span>

<span data-ttu-id="a6c49-421">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-422">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-422">Compose mode</span></span>

<span data-ttu-id="a6c49-423">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-424">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-424">Type:</span></span>

*   <span data-ttu-id="a6c49-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-426">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-426">Requirements</span></span>

|<span data-ttu-id="a6c49-427">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-427">Requirement</span></span>| <span data-ttu-id="a6c49-428">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-429">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-430">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-430">1.0</span></span>|
|[<span data-ttu-id="a6c49-431">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-432">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-433">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-434">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-435">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="a6c49-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="a6c49-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="a6c49-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="a6c49-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-441">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-442">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-442">Type:</span></span>

*   [<span data-ttu-id="a6c49-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="a6c49-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="a6c49-444">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-444">Requirements</span></span>

|<span data-ttu-id="a6c49-445">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-445">Requirement</span></span>| <span data-ttu-id="a6c49-446">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-448">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-448">1.0</span></span>|
|[<span data-ttu-id="a6c49-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-450">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-453">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="a6c49-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6c49-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="a6c49-455">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="a6c49-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-458">Read mode</span></span>

<span data-ttu-id="a6c49-459">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-460">Compose mode</span></span>

<span data-ttu-id="a6c49-461">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="a6c49-462">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-463">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-463">Type:</span></span>

*   <span data-ttu-id="a6c49-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="a6c49-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-465">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-465">Requirements</span></span>

|<span data-ttu-id="a6c49-466">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-466">Requirement</span></span>| <span data-ttu-id="a6c49-467">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-468">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-469">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-469">1.0</span></span>|
|[<span data-ttu-id="a6c49-470">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-471">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-472">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-473">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-474">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-474">Example</span></span>

<span data-ttu-id="a6c49-475">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="a6c49-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a6c49-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="a6c49-477">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="a6c49-478">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-479">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-479">Read mode</span></span>

<span data-ttu-id="a6c49-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-482">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-482">Compose mode</span></span>

<span data-ttu-id="a6c49-483">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="a6c49-484">型:</span><span class="sxs-lookup"><span data-stu-id="a6c49-484">Type:</span></span>

*   <span data-ttu-id="a6c49-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="a6c49-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-486">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-486">Requirements</span></span>

|<span data-ttu-id="a6c49-487">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-487">Requirement</span></span>| <span data-ttu-id="a6c49-488">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-489">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-490">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-490">1.0</span></span>|
|[<span data-ttu-id="a6c49-491">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-492">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-493">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-494">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="a6c49-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="a6c49-496">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="a6c49-497">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="a6c49-498">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-498">Read mode</span></span>

<span data-ttu-id="a6c49-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="a6c49-501">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="a6c49-501">Compose mode</span></span>

<span data-ttu-id="a6c49-502">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="a6c49-503">種類:</span><span class="sxs-lookup"><span data-stu-id="a6c49-503">Type:</span></span>

*   <span data-ttu-id="a6c49-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="a6c49-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-505">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-505">Requirements</span></span>

|<span data-ttu-id="a6c49-506">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-506">Requirement</span></span>| <span data-ttu-id="a6c49-507">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-509">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-509">1.0</span></span>|
|[<span data-ttu-id="a6c49-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-511">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-513">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-514">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="a6c49-515">メソッド</span><span class="sxs-lookup"><span data-stu-id="a6c49-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="a6c49-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6c49-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a6c49-517">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="a6c49-518">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="a6c49-519">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-520">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-520">Parameters:</span></span>

|<span data-ttu-id="a6c49-521">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-521">Name</span></span>| <span data-ttu-id="a6c49-522">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-522">Type</span></span>| <span data-ttu-id="a6c49-523">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-523">Attributes</span></span>| <span data-ttu-id="a6c49-524">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="a6c49-525">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-525">String</span></span>||<span data-ttu-id="a6c49-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a6c49-528">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-528">String</span></span>||<span data-ttu-id="a6c49-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a6c49-531">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-531">Object</span></span>| <span data-ttu-id="a6c49-532">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-532">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-533">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a6c49-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a6c49-534">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-534">Object</span></span>| <span data-ttu-id="a6c49-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-535">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-536">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a6c49-537">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-537">function</span></span>| <span data-ttu-id="a6c49-538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-538">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-539">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6c49-540">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a6c49-541">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6c49-542">エラー</span><span class="sxs-lookup"><span data-stu-id="a6c49-542">Errors</span></span>

| <span data-ttu-id="a6c49-543">エラー コード</span><span class="sxs-lookup"><span data-stu-id="a6c49-543">Error code</span></span> | <span data-ttu-id="a6c49-544">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="a6c49-545">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="a6c49-546">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a6c49-547">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-548">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-548">Requirements</span></span>

|<span data-ttu-id="a6c49-549">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-549">Requirement</span></span>| <span data-ttu-id="a6c49-550">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-551">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-552">1.1</span><span class="sxs-lookup"><span data-stu-id="a6c49-552">1.1</span></span>|
|[<span data-ttu-id="a6c49-553">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6c49-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-556">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-557">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-557">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="a6c49-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6c49-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="a6c49-559">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="a6c49-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="a6c49-563">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="a6c49-564">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-565">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-565">Parameters:</span></span>

|<span data-ttu-id="a6c49-566">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-566">Name</span></span>| <span data-ttu-id="a6c49-567">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-567">Type</span></span>| <span data-ttu-id="a6c49-568">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-568">Attributes</span></span>| <span data-ttu-id="a6c49-569">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="a6c49-570">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-570">String</span></span>||<span data-ttu-id="a6c49-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="a6c49-573">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-573">String</span></span>||<span data-ttu-id="a6c49-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="a6c49-576">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-576">Object</span></span>| <span data-ttu-id="a6c49-577">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-577">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-578">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a6c49-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a6c49-579">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-579">Object</span></span>| <span data-ttu-id="a6c49-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-580">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-581">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a6c49-582">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-582">function</span></span>| <span data-ttu-id="a6c49-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-583">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-584">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6c49-585">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="a6c49-586">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6c49-587">エラー</span><span class="sxs-lookup"><span data-stu-id="a6c49-587">Errors</span></span>

| <span data-ttu-id="a6c49-588">エラー コード</span><span class="sxs-lookup"><span data-stu-id="a6c49-588">Error code</span></span> | <span data-ttu-id="a6c49-589">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="a6c49-590">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-591">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-591">Requirements</span></span>

|<span data-ttu-id="a6c49-592">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-592">Requirement</span></span>| <span data-ttu-id="a6c49-593">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-594">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-595">1.1</span><span class="sxs-lookup"><span data-stu-id="a6c49-595">1.1</span></span>|
|[<span data-ttu-id="a6c49-596">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6c49-598">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-599">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-600">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-600">Example</span></span>

<span data-ttu-id="a6c49-601">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="a6c49-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a6c49-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="a6c49-603">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-604">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6c49-605">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a6c49-606">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="a6c49-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="a6c49-p137">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-610">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-610">Parameters:</span></span>

|<span data-ttu-id="a6c49-611">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-611">Name</span></span>| <span data-ttu-id="a6c49-612">種類</span><span class="sxs-lookup"><span data-stu-id="a6c49-612">Type</span></span>| <span data-ttu-id="a6c49-613">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a6c49-614">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-614">String &#124; Object</span></span>| |<span data-ttu-id="a6c49-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a6c49-617">**または**</span><span class="sxs-lookup"><span data-stu-id="a6c49-617">**OR**</span></span><br/><span data-ttu-id="a6c49-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a6c49-620">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-620">String</span></span> | <span data-ttu-id="a6c49-621">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-621">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a6c49-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a6c49-625">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-625">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-626">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="a6c49-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a6c49-627">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-627">String</span></span> | | <span data-ttu-id="a6c49-p141">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a6c49-630">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-630">String</span></span> | | <span data-ttu-id="a6c49-631">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a6c49-632">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-632">String</span></span> | | <span data-ttu-id="a6c49-p142">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a6c49-635">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-635">String</span></span> | | <span data-ttu-id="a6c49-p143">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a6c49-639">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-639">function</span></span> | <span data-ttu-id="a6c49-640">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-640">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-641">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-642">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-642">Requirements</span></span>

|<span data-ttu-id="a6c49-643">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-643">Requirement</span></span>| <span data-ttu-id="a6c49-644">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-645">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-646">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-646">1.0</span></span>|
|[<span data-ttu-id="a6c49-647">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-648">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-649">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-650">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6c49-651">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-651">Examples</span></span>

<span data-ttu-id="a6c49-652">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="a6c49-653">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="a6c49-654">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a6c49-655">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-655">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="a6c49-656">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-656">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="a6c49-657">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="a6c49-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="a6c49-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="a6c49-659">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-660">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6c49-661">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="a6c49-662">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="a6c49-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="a6c49-p144">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-666">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-666">Parameters:</span></span>

|<span data-ttu-id="a6c49-667">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-667">Name</span></span>| <span data-ttu-id="a6c49-668">種類</span><span class="sxs-lookup"><span data-stu-id="a6c49-668">Type</span></span>| <span data-ttu-id="a6c49-669">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="a6c49-670">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-670">String &#124; Object</span></span>| | <span data-ttu-id="a6c49-p145">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="a6c49-673">**または**</span><span class="sxs-lookup"><span data-stu-id="a6c49-673">**OR**</span></span><br/><span data-ttu-id="a6c49-p146">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="a6c49-676">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-676">String</span></span> | <span data-ttu-id="a6c49-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-677">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-p147">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="a6c49-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="a6c49-681">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-681">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-682">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="a6c49-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="a6c49-683">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-683">String</span></span> | | <span data-ttu-id="a6c49-p148">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="a6c49-686">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-686">String</span></span> | | <span data-ttu-id="a6c49-687">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="a6c49-688">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-688">String</span></span> | | <span data-ttu-id="a6c49-p149">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="a6c49-691">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-691">String</span></span> | | <span data-ttu-id="a6c49-p150">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="a6c49-695">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-695">function</span></span> | <span data-ttu-id="a6c49-696">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-696">&lt;optional&gt;</span></span> | <span data-ttu-id="a6c49-697">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-698">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-698">Requirements</span></span>

|<span data-ttu-id="a6c49-699">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-699">Requirement</span></span>| <span data-ttu-id="a6c49-700">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-701">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-702">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-702">1.0</span></span>|
|[<span data-ttu-id="a6c49-703">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-704">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-705">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-706">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="a6c49-707">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-707">Examples</span></span>

<span data-ttu-id="a6c49-708">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="a6c49-709">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="a6c49-710">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="a6c49-711">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-711">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="a6c49-712">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-712">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="a6c49-713">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="a6c49-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="a6c49-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="a6c49-715">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-716">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-717">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-717">Requirements</span></span>

|<span data-ttu-id="a6c49-718">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-718">Requirement</span></span>| <span data-ttu-id="a6c49-719">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-720">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-721">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-721">1.0</span></span>|
|[<span data-ttu-id="a6c49-722">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-723">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-724">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-725">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-726">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-726">Returns:</span></span>

<span data-ttu-id="a6c49-727">型:[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="a6c49-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="a6c49-728">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-728">Example</span></span>

<span data-ttu-id="a6c49-729">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="a6c49-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="a6c49-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a6c49-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a6c49-731">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-732">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-733">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-733">Parameters:</span></span>

|<span data-ttu-id="a6c49-734">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-734">Name</span></span>| <span data-ttu-id="a6c49-735">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-735">Type</span></span>| <span data-ttu-id="a6c49-736">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="a6c49-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="a6c49-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="a6c49-738">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="a6c49-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6c49-739">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6c49-739">Requirements</span></span>

|<span data-ttu-id="a6c49-740">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-740">Requirement</span></span>| <span data-ttu-id="a6c49-741">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-742">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-743">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-743">1.0</span></span>|
|[<span data-ttu-id="a6c49-744">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-745">制限あり</span><span class="sxs-lookup"><span data-stu-id="a6c49-745">Restricted</span></span>|
|[<span data-ttu-id="a6c49-746">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-747">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-748">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-748">Returns:</span></span>

<span data-ttu-id="a6c49-749">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="a6c49-750">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="a6c49-751">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="a6c49-752">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="a6c49-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="a6c49-753">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="a6c49-753">Value of `entityType`</span></span> | <span data-ttu-id="a6c49-754">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="a6c49-754">Type of objects in returned array</span></span> | <span data-ttu-id="a6c49-755">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="a6c49-756">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-756">String</span></span> | <span data-ttu-id="a6c49-757">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="a6c49-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="a6c49-758">連絡先</span><span class="sxs-lookup"><span data-stu-id="a6c49-758">Contact</span></span> | <span data-ttu-id="a6c49-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6c49-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="a6c49-760">文字列</span><span class="sxs-lookup"><span data-stu-id="a6c49-760">String</span></span> | <span data-ttu-id="a6c49-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6c49-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="a6c49-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="a6c49-762">MeetingSuggestion</span></span> | <span data-ttu-id="a6c49-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6c49-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="a6c49-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="a6c49-764">PhoneNumber</span></span> | <span data-ttu-id="a6c49-765">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="a6c49-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="a6c49-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="a6c49-766">TaskSuggestion</span></span> | <span data-ttu-id="a6c49-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="a6c49-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="a6c49-768">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-768">String</span></span> | <span data-ttu-id="a6c49-769">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="a6c49-769">**Restricted**</span></span> |

<span data-ttu-id="a6c49-770">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a6c49-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="a6c49-771">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-771">Example</span></span>

<span data-ttu-id="a6c49-772">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="a6c49-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="a6c49-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="a6c49-774">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-775">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6c49-776">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-777">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-777">Parameters:</span></span>

|<span data-ttu-id="a6c49-778">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-778">Name</span></span>| <span data-ttu-id="a6c49-779">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-779">Type</span></span>| <span data-ttu-id="a6c49-780">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a6c49-781">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-781">String</span></span>|<span data-ttu-id="a6c49-782">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="a6c49-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6c49-783">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-783">Requirements</span></span>

|<span data-ttu-id="a6c49-784">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-784">Requirement</span></span>| <span data-ttu-id="a6c49-785">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-786">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-787">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-787">1.0</span></span>|
|[<span data-ttu-id="a6c49-788">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-789">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-790">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-791">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-792">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-792">Returns:</span></span>

<span data-ttu-id="a6c49-p152">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="a6c49-795">型:Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="a6c49-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="a6c49-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="a6c49-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="a6c49-797">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-798">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6c49-p153">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="a6c49-802">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="a6c49-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="a6c49-803">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="a6c49-p154">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6c49-806">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-806">Requirements</span></span>

|<span data-ttu-id="a6c49-807">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-807">Requirement</span></span>| <span data-ttu-id="a6c49-808">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-809">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-810">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-810">1.0</span></span>|
|[<span data-ttu-id="a6c49-811">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-812">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-813">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-814">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-815">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-815">Returns:</span></span>

<span data-ttu-id="a6c49-p155">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="a6c49-818">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6c49-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6c49-819">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="a6c49-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6c49-820">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-820">Example</span></span>

<span data-ttu-id="a6c49-821">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="a6c49-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="a6c49-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="a6c49-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="a6c49-823">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="a6c49-824">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="a6c49-825">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="a6c49-p156">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-828">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-828">Parameters:</span></span>

|<span data-ttu-id="a6c49-829">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-829">Name</span></span>| <span data-ttu-id="a6c49-830">種類</span><span class="sxs-lookup"><span data-stu-id="a6c49-830">Type</span></span>| <span data-ttu-id="a6c49-831">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="a6c49-832">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-832">String</span></span>|<span data-ttu-id="a6c49-833">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="a6c49-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6c49-834">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-834">Requirements</span></span>

|<span data-ttu-id="a6c49-835">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-835">Requirement</span></span>| <span data-ttu-id="a6c49-836">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-837">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-838">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-838">1.0</span></span>|
|[<span data-ttu-id="a6c49-839">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-840">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-841">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-842">読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-843">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-843">Returns:</span></span>

<span data-ttu-id="a6c49-844">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="a6c49-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="a6c49-845">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6c49-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6c49-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="a6c49-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6c49-847">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="a6c49-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="a6c49-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="a6c49-849">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="a6c49-p157">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-852">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-852">Parameters:</span></span>

|<span data-ttu-id="a6c49-853">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-853">Name</span></span>| <span data-ttu-id="a6c49-854">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-854">Type</span></span>| <span data-ttu-id="a6c49-855">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-855">Attributes</span></span>| <span data-ttu-id="a6c49-856">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="a6c49-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a6c49-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="a6c49-p158">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="a6c49-861">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-861">Object</span></span>| <span data-ttu-id="a6c49-862">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-862">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-863">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a6c49-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a6c49-864">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-864">Object</span></span>| <span data-ttu-id="a6c49-865">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-865">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-866">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a6c49-867">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-867">function</span></span>||<span data-ttu-id="a6c49-868">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6c49-869">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="a6c49-870">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="a6c49-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6c49-871">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-871">Requirements</span></span>

|<span data-ttu-id="a6c49-872">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-872">Requirement</span></span>| <span data-ttu-id="a6c49-873">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-874">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-875">1.2</span><span class="sxs-lookup"><span data-stu-id="a6c49-875">1.2</span></span>|
|[<span data-ttu-id="a6c49-876">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6c49-878">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-879">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6c49-880">戻り値:</span><span class="sxs-lookup"><span data-stu-id="a6c49-880">Returns:</span></span>

<span data-ttu-id="a6c49-881">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="a6c49-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="a6c49-882">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="a6c49-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="a6c49-883">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="a6c49-884">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-884">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="a6c49-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a6c49-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="a6c49-886">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="a6c49-p160">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-890">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-890">Parameters:</span></span>

|<span data-ttu-id="a6c49-891">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-891">Name</span></span>| <span data-ttu-id="a6c49-892">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-892">Type</span></span>| <span data-ttu-id="a6c49-893">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-893">Attributes</span></span>| <span data-ttu-id="a6c49-894">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a6c49-895">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-895">function</span></span>||<span data-ttu-id="a6c49-896">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6c49-897">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="a6c49-898">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="a6c49-899">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-899">Object</span></span>| <span data-ttu-id="a6c49-900">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-900">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-901">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="a6c49-902">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6c49-903">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-903">Requirements</span></span>

|<span data-ttu-id="a6c49-904">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-904">Requirement</span></span>| <span data-ttu-id="a6c49-905">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-906">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-907">1.0</span><span class="sxs-lookup"><span data-stu-id="a6c49-907">1.0</span></span>|
|[<span data-ttu-id="a6c49-908">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-909">ReadItem</span></span>|
|[<span data-ttu-id="a6c49-910">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-911">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="a6c49-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-912">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-912">Example</span></span>

<span data-ttu-id="a6c49-p163">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="a6c49-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="a6c49-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="a6c49-917">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="a6c49-p164">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-922">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-922">Parameters:</span></span>

|<span data-ttu-id="a6c49-923">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-923">Name</span></span>| <span data-ttu-id="a6c49-924">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-924">Type</span></span>| <span data-ttu-id="a6c49-925">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-925">Attributes</span></span>| <span data-ttu-id="a6c49-926">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="a6c49-927">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-927">String</span></span>||<span data-ttu-id="a6c49-928">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="a6c49-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="a6c49-929">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="a6c49-929">Object</span></span>| <span data-ttu-id="a6c49-930">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-930">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-931">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a6c49-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a6c49-932">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-932">Object</span></span>| <span data-ttu-id="a6c49-933">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-933">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-934">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="a6c49-935">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-935">function</span></span>| <span data-ttu-id="a6c49-936">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-936">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-937">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="a6c49-938">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6c49-939">エラー</span><span class="sxs-lookup"><span data-stu-id="a6c49-939">Errors</span></span>

| <span data-ttu-id="a6c49-940">エラー コード</span><span class="sxs-lookup"><span data-stu-id="a6c49-940">Error code</span></span> | <span data-ttu-id="a6c49-941">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="a6c49-942">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="a6c49-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-943">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-943">Requirements</span></span>

|<span data-ttu-id="a6c49-944">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-944">Requirement</span></span>| <span data-ttu-id="a6c49-945">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-946">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-947">1.1</span><span class="sxs-lookup"><span data-stu-id="a6c49-947">1.1</span></span>|
|[<span data-ttu-id="a6c49-948">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6c49-950">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-951">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-952">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-952">Example</span></span>

<span data-ttu-id="a6c49-953">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="a6c49-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="a6c49-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="a6c49-955">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="a6c49-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="a6c49-p165">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6c49-959">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="a6c49-959">Parameters:</span></span>

|<span data-ttu-id="a6c49-960">名前</span><span class="sxs-lookup"><span data-stu-id="a6c49-960">Name</span></span>| <span data-ttu-id="a6c49-961">型</span><span class="sxs-lookup"><span data-stu-id="a6c49-961">Type</span></span>| <span data-ttu-id="a6c49-962">属性</span><span class="sxs-lookup"><span data-stu-id="a6c49-962">Attributes</span></span>| <span data-ttu-id="a6c49-963">説明</span><span class="sxs-lookup"><span data-stu-id="a6c49-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a6c49-964">String</span><span class="sxs-lookup"><span data-stu-id="a6c49-964">String</span></span>||<span data-ttu-id="a6c49-p166">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="a6c49-968">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-968">Object</span></span>| <span data-ttu-id="a6c49-969">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-969">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-970">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="a6c49-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="a6c49-971">Object</span><span class="sxs-lookup"><span data-stu-id="a6c49-971">Object</span></span>| <span data-ttu-id="a6c49-972">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-972">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-973">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="a6c49-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="a6c49-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="a6c49-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6c49-975">&lt;optional&gt;</span></span>|<span data-ttu-id="a6c49-p167">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="a6c49-p168">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="a6c49-980">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="a6c49-981">function</span><span class="sxs-lookup"><span data-stu-id="a6c49-981">function</span></span>||<span data-ttu-id="a6c49-982">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="a6c49-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6c49-983">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-983">Requirements</span></span>

|<span data-ttu-id="a6c49-984">要件</span><span class="sxs-lookup"><span data-stu-id="a6c49-984">Requirement</span></span>| <span data-ttu-id="a6c49-985">値</span><span class="sxs-lookup"><span data-stu-id="a6c49-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6c49-986">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="a6c49-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6c49-987">1.2</span><span class="sxs-lookup"><span data-stu-id="a6c49-987">1.2</span></span>|
|[<span data-ttu-id="a6c49-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="a6c49-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6c49-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="a6c49-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="a6c49-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="a6c49-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a6c49-991">作成</span><span class="sxs-lookup"><span data-stu-id="a6c49-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="a6c49-992">例</span><span class="sxs-lookup"><span data-stu-id="a6c49-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
