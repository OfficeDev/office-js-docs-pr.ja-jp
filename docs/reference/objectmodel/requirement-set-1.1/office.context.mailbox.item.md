---
title: Office.context.mailbox.item の要件は、1.1 を設定
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: ce8c10987c08609eba90a3a957b372114e62cd81
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701877"
---
# <a name="item"></a><span data-ttu-id="e5c5c-102">item</span><span class="sxs-lookup"><span data-stu-id="e5c5c-102">item</span></span>

### <span data-ttu-id="e5c5c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="e5c5c-p102">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-107">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-107">Requirements</span></span>

|<span data-ttu-id="e5c5c-108">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-108">Requirement</span></span>| <span data-ttu-id="e5c5c-109">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-111">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-111">1.0</span></span>|
|[<span data-ttu-id="e5c5c-112">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-113">制限あり</span><span class="sxs-lookup"><span data-stu-id="e5c5c-113">Restricted</span></span>|
|[<span data-ttu-id="e5c5c-114">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-115">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="e5c5c-116">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-116">Example</span></span>

<span data-ttu-id="e5c5c-117">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e5c5c-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="e5c5c-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="e5c5c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c5c-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="e5c5c-p103">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-122">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e5c5c-123">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-124">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-124">Type:</span></span>

*   <span data-ttu-id="e5c5c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e5c5c-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-126">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-126">Requirements</span></span>

|<span data-ttu-id="e5c5c-127">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-127">Requirement</span></span>| <span data-ttu-id="e5c5c-128">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-129">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-130">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-130">1.0</span></span>|
|[<span data-ttu-id="e5c5c-131">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-132">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-134">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-135">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-135">Example</span></span>

<span data-ttu-id="e5c5c-136">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="e5c5c-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="e5c5c-138">メッセージの BCC (ブラインド カーボン コピー) 行の受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e5c5c-139">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-140">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-140">Type:</span></span>

*   [<span data-ttu-id="e5c5c-141">Recipients</span><span class="sxs-lookup"><span data-stu-id="e5c5c-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-142">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-142">Requirements</span></span>

|<span data-ttu-id="e5c5c-143">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-143">Requirement</span></span>| <span data-ttu-id="e5c5c-144">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-145">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-146">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c5c-146">1.1</span></span>|
|[<span data-ttu-id="e5c5c-147">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-148">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-149">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-150">作成</span><span class="sxs-lookup"><span data-stu-id="e5c5c-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-151">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="e5c5c-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="e5c5c-153">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-154">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-154">Type:</span></span>

*   [<span data-ttu-id="e5c5c-155">Body</span><span class="sxs-lookup"><span data-stu-id="e5c5c-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-156">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-156">Requirements</span></span>

|<span data-ttu-id="e5c5c-157">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-157">Requirement</span></span>| <span data-ttu-id="e5c5c-158">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-159">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-160">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c5c-160">1.1</span></span>|
|[<span data-ttu-id="e5c5c-161">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-162">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-163">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-164">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="e5c5c-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="e5c5c-166">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e5c5c-167">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-168">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-168">Read mode</span></span>

<span data-ttu-id="e5c5c-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-171">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-171">Compose mode</span></span>

<span data-ttu-id="e5c5c-172">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-173">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-173">Type:</span></span>

*   <span data-ttu-id="e5c5c-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-175">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-175">Requirements</span></span>

|<span data-ttu-id="e5c5c-176">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-176">Requirement</span></span>| <span data-ttu-id="e5c5c-177">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-178">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-179">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-179">1.0</span></span>|
|[<span data-ttu-id="e5c5c-180">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-181">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-183">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-184">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="e5c5c-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="e5c5c-186">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e5c5c-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e5c5c-p109">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-191">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-191">Type:</span></span>

*   <span data-ttu-id="e5c5c-192">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-193">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-193">Requirements</span></span>

|<span data-ttu-id="e5c5c-194">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-194">Requirement</span></span>| <span data-ttu-id="e5c5c-195">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-196">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-197">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-197">1.0</span></span>|
|[<span data-ttu-id="e5c5c-198">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-199">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e5c5c-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="e5c5c-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="e5c5c-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="e5c5c-p110">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-205">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-205">Type:</span></span>

*   <span data-ttu-id="e5c5c-206">日付</span><span class="sxs-lookup"><span data-stu-id="e5c5c-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-207">Requirements</span></span>

|<span data-ttu-id="e5c5c-208">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-208">Requirement</span></span>| <span data-ttu-id="e5c5c-209">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-211">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-211">1.0</span></span>|
|[<span data-ttu-id="e5c5c-212">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-213">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-215">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-216">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="e5c5c-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="e5c5c-217">dateTimeModified :Date</span></span>

<span data-ttu-id="e5c5c-p111">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-220">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-221">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-221">Type:</span></span>

*   <span data-ttu-id="e5c5c-222">日付</span><span class="sxs-lookup"><span data-stu-id="e5c5c-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-223">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-223">Requirements</span></span>

|<span data-ttu-id="e5c5c-224">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-224">Requirement</span></span>| <span data-ttu-id="e5c5c-225">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-226">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-227">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-227">1.0</span></span>|
|[<span data-ttu-id="e5c5c-228">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-229">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-230">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-231">Read</span><span class="sxs-lookup"><span data-stu-id="e5c5c-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-232">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="e5c5c-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="e5c5c-234">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e5c5c-p112">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-237">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-237">Read mode</span></span>

<span data-ttu-id="e5c5c-238">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-239">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-239">Compose mode</span></span>

<span data-ttu-id="e5c5c-240">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e5c5c-241">[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-242">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-242">Type:</span></span>

*   <span data-ttu-id="e5c5c-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-244">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-244">Requirements</span></span>

|<span data-ttu-id="e5c5c-245">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-245">Requirement</span></span>| <span data-ttu-id="e5c5c-246">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-247">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-248">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-248">1.0</span></span>|
|[<span data-ttu-id="e5c5c-249">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-250">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-251">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-252">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-253">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-253">Example</span></span>

<span data-ttu-id="e5c5c-254">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="e5c5c-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5c5c-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="e5c5c-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-260">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-261">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-261">Type:</span></span>

*   [<span data-ttu-id="e5c5c-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5c5c-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-263">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-263">Requirements</span></span>

|<span data-ttu-id="e5c5c-264">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-264">Requirement</span></span>| <span data-ttu-id="e5c5c-265">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-266">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-267">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-267">1.0</span></span>|
|[<span data-ttu-id="e5c5c-268">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-269">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-270">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-271">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="e5c5c-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-272">internetMessageId :String</span></span>

<span data-ttu-id="e5c5c-p115">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-275">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-275">Type:</span></span>

*   <span data-ttu-id="e5c5c-276">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-277">Requirements</span></span>

|<span data-ttu-id="e5c5c-278">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-278">Requirement</span></span>| <span data-ttu-id="e5c5c-279">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-281">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-281">1.0</span></span>|
|[<span data-ttu-id="e5c5c-282">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-283">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-285">Read</span><span class="sxs-lookup"><span data-stu-id="e5c5c-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-286">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="e5c5c-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-287">itemClass :String</span></span>

<span data-ttu-id="e5c5c-p116">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e5c5c-p117">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="e5c5c-292">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-292">Type</span></span> | <span data-ttu-id="e5c5c-293">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-293">Description</span></span> | <span data-ttu-id="e5c5c-294">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="e5c5c-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="e5c5c-295">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="e5c5c-295">Appointment items</span></span> | <span data-ttu-id="e5c5c-296">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurrence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="e5c5c-297">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="e5c5c-297">Message items</span></span> | <span data-ttu-id="e5c5c-298">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="e5c5c-299">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-300">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-300">Type:</span></span>

*   <span data-ttu-id="e5c5c-301">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-302">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-302">Requirements</span></span>

|<span data-ttu-id="e5c5c-303">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-303">Requirement</span></span>| <span data-ttu-id="e5c5c-304">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-305">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-306">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-306">1.0</span></span>|
|[<span data-ttu-id="e5c5c-307">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-308">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-309">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-310">Read</span><span class="sxs-lookup"><span data-stu-id="e5c5c-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-311">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e5c5c-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-312">(nullable) itemId :String</span></span>

<span data-ttu-id="e5c5c-p118">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-315">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e5c5c-316">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e5c5c-317">この値を使用して REST API を呼び出す前に、要件セット 1.3 から使用できる `Office.context.mailbox.convertToRestId` を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="e5c5c-318">詳細は、「[Outlook アドインからの Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-319">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-319">Type:</span></span>

*   <span data-ttu-id="e5c5c-320">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-321">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-321">Requirements</span></span>

|<span data-ttu-id="e5c5c-322">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-322">Requirement</span></span>| <span data-ttu-id="e5c5c-323">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-325">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-325">1.0</span></span>|
|[<span data-ttu-id="e5c5c-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-327">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-329">Read</span><span class="sxs-lookup"><span data-stu-id="e5c5c-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-330">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-330">Example</span></span>

<span data-ttu-id="e5c5c-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="e5c5c-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e5c5c-334">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e5c5c-335">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-336">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-336">Type:</span></span>

*   [<span data-ttu-id="e5c5c-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e5c5c-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-338">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-338">Requirements</span></span>

|<span data-ttu-id="e5c5c-339">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-339">Requirement</span></span>| <span data-ttu-id="e5c5c-340">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-342">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-342">1.0</span></span>|
|[<span data-ttu-id="e5c5c-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-344">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-346">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-347">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="e5c5c-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="e5c5c-349">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-350">Read mode</span></span>

<span data-ttu-id="e5c5c-351">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-352">Compose mode</span></span>

<span data-ttu-id="e5c5c-353">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-354">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-354">Type:</span></span>

*   <span data-ttu-id="e5c5c-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-356">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-356">Requirements</span></span>

|<span data-ttu-id="e5c5c-357">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-357">Requirement</span></span>| <span data-ttu-id="e5c5c-358">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-360">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-360">1.0</span></span>|
|[<span data-ttu-id="e5c5c-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-362">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-364">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-365">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e5c5c-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-366">normalizedSubject :String</span></span>

<span data-ttu-id="e5c5c-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e5c5c-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-371">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-371">Type:</span></span>

*   <span data-ttu-id="e5c5c-372">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-373">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-373">Requirements</span></span>

|<span data-ttu-id="e5c5c-374">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-374">Requirement</span></span>| <span data-ttu-id="e5c5c-375">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-376">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-377">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-377">1.0</span></span>|
|[<span data-ttu-id="e5c5c-378">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-379">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-381">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-382">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="e5c5c-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="e5c5c-384">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e5c5c-385">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-386">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-386">Read mode</span></span>

<span data-ttu-id="e5c5c-387">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-388">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-388">Compose mode</span></span>

<span data-ttu-id="e5c5c-389">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-390">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-390">Type:</span></span>

*   <span data-ttu-id="e5c5c-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-392">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-392">Requirements</span></span>

|<span data-ttu-id="e5c5c-393">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-393">Requirement</span></span>| <span data-ttu-id="e5c5c-394">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-395">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-396">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-396">1.0</span></span>|
|[<span data-ttu-id="e5c5c-397">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-398">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-399">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-400">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-401">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="e5c5c-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5c5c-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-405">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-405">Type:</span></span>

*   [<span data-ttu-id="e5c5c-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5c5c-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-407">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-407">Requirements</span></span>

|<span data-ttu-id="e5c5c-408">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-408">Requirement</span></span>| <span data-ttu-id="e5c5c-409">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-410">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-411">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-411">1.0</span></span>|
|[<span data-ttu-id="e5c5c-412">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-413">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-414">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-415">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-416">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="e5c5c-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="e5c5c-418">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e5c5c-419">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-420">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-420">Read mode</span></span>

<span data-ttu-id="e5c5c-421">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-422">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-422">Compose mode</span></span>

<span data-ttu-id="e5c5c-423">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-424">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-424">Type:</span></span>

*   <span data-ttu-id="e5c5c-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-426">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-426">Requirements</span></span>

|<span data-ttu-id="e5c5c-427">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-427">Requirement</span></span>| <span data-ttu-id="e5c5c-428">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-429">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-430">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-430">1.0</span></span>|
|[<span data-ttu-id="e5c5c-431">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-432">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-433">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-434">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-435">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="e5c5c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="e5c5c-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e5c5c-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-441">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-442">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-442">Type:</span></span>

*   [<span data-ttu-id="e5c5c-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e5c5c-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e5c5c-444">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-444">Requirements</span></span>

|<span data-ttu-id="e5c5c-445">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-445">Requirement</span></span>| <span data-ttu-id="e5c5c-446">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-447">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-448">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-448">1.0</span></span>|
|[<span data-ttu-id="e5c5c-449">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-450">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-451">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-452">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-453">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="e5c5c-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="e5c5c-455">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e5c5c-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-458">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-458">Read mode</span></span>

<span data-ttu-id="e5c5c-459">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-460">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-460">Compose mode</span></span>

<span data-ttu-id="e5c5c-461">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e5c5c-462">[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-463">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-463">Type:</span></span>

*   <span data-ttu-id="e5c5c-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-465">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-465">Requirements</span></span>

|<span data-ttu-id="e5c5c-466">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-466">Requirement</span></span>| <span data-ttu-id="e5c5c-467">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-468">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-469">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-469">1.0</span></span>|
|[<span data-ttu-id="e5c5c-470">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-471">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-472">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-473">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-474">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-474">Example</span></span>

<span data-ttu-id="e5c5c-475">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="e5c5c-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="e5c5c-477">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e5c5c-478">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-479">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-479">Read mode</span></span>

<span data-ttu-id="e5c5c-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-482">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-482">Compose mode</span></span>

<span data-ttu-id="e5c5c-483">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e5c5c-484">型:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-484">Type:</span></span>

*   <span data-ttu-id="e5c5c-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-486">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-486">Requirements</span></span>

|<span data-ttu-id="e5c5c-487">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-487">Requirement</span></span>| <span data-ttu-id="e5c5c-488">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-489">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-490">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-490">1.0</span></span>|
|[<span data-ttu-id="e5c5c-491">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-492">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-493">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-494">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="e5c5c-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="e5c5c-496">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e5c5c-497">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e5c5c-498">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-498">Read mode</span></span>

<span data-ttu-id="e5c5c-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e5c5c-501">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-501">Compose mode</span></span>

<span data-ttu-id="e5c5c-502">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e5c5c-503">種類:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-503">Type:</span></span>

*   <span data-ttu-id="e5c5c-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-505">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-505">Requirements</span></span>

|<span data-ttu-id="e5c5c-506">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-506">Requirement</span></span>| <span data-ttu-id="e5c5c-507">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-508">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-509">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-509">1.0</span></span>|
|[<span data-ttu-id="e5c5c-510">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-511">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-512">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-513">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-514">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="e5c5c-515">メソッド</span><span class="sxs-lookup"><span data-stu-id="e5c5c-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e5c5c-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c5c-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5c5c-517">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e5c5c-518">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e5c5c-519">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-520">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-520">Parameters:</span></span>

|<span data-ttu-id="e5c5c-521">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-521">Name</span></span>| <span data-ttu-id="e5c5c-522">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-522">Type</span></span>| <span data-ttu-id="e5c5c-523">属性</span><span class="sxs-lookup"><span data-stu-id="e5c5c-523">Attributes</span></span>| <span data-ttu-id="e5c5c-524">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="e5c5c-525">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-525">String</span></span>||<span data-ttu-id="e5c5c-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e5c5c-528">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-528">String</span></span>||<span data-ttu-id="e5c5c-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e5c5c-531">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-531">Object</span></span>| <span data-ttu-id="e5c5c-532">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-532">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-533">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5c5c-534">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-534">Object</span></span>| <span data-ttu-id="e5c5c-535">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-535">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-536">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5c5c-537">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-537">function</span></span>| <span data-ttu-id="e5c5c-538">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-538">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-539">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c5c-540">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5c5c-541">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c5c-542">エラー</span><span class="sxs-lookup"><span data-stu-id="e5c5c-542">Errors</span></span>

| <span data-ttu-id="e5c5c-543">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-543">Error code</span></span> | <span data-ttu-id="e5c5c-544">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="e5c5c-545">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="e5c5c-546">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e5c5c-547">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5c5c-548">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-548">Requirements</span></span>

|<span data-ttu-id="e5c5c-549">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-549">Requirement</span></span>| <span data-ttu-id="e5c5c-550">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-551">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-552">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c5c-552">1.1</span></span>|
|[<span data-ttu-id="e5c5c-553">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c5c-555">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-556">作成</span><span class="sxs-lookup"><span data-stu-id="e5c5c-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-557">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e5c5c-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c5c-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e5c5c-559">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e5c5c-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e5c5c-563">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e5c5c-564">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-565">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-565">Parameters:</span></span>

|<span data-ttu-id="e5c5c-566">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-566">Name</span></span>| <span data-ttu-id="e5c5c-567">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-567">Type</span></span>| <span data-ttu-id="e5c5c-568">属性</span><span class="sxs-lookup"><span data-stu-id="e5c5c-568">Attributes</span></span>| <span data-ttu-id="e5c5c-569">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="e5c5c-570">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-570">String</span></span>||<span data-ttu-id="e5c5c-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="e5c5c-573">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-573">String</span></span>||<span data-ttu-id="e5c5c-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="e5c5c-576">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-576">Object</span></span>| <span data-ttu-id="e5c5c-577">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-577">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-578">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5c5c-579">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-579">Object</span></span>| <span data-ttu-id="e5c5c-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-580">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-581">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5c5c-582">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-582">function</span></span>| <span data-ttu-id="e5c5c-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-583">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-584">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c5c-585">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e5c5c-586">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c5c-587">エラー</span><span class="sxs-lookup"><span data-stu-id="e5c5c-587">Errors</span></span>

| <span data-ttu-id="e5c5c-588">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-588">Error code</span></span> | <span data-ttu-id="e5c5c-589">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="e5c5c-590">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5c5c-591">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-591">Requirements</span></span>

|<span data-ttu-id="e5c5c-592">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-592">Requirement</span></span>| <span data-ttu-id="e5c5c-593">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-594">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-595">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c5c-595">1.1</span></span>|
|[<span data-ttu-id="e5c5c-596">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c5c-598">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-599">作成</span><span class="sxs-lookup"><span data-stu-id="e5c5c-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-600">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-600">Example</span></span>

<span data-ttu-id="e5c5c-601">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="e5c5c-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="e5c5c-603">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-604">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5c5c-605">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5c5c-606">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-607">`displayReplyAllForm` に対する呼び出しに添付ファイルを含める機能は、要件セット 1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="e5c5c-608">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyAllForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-609">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-609">Parameters:</span></span>

|<span data-ttu-id="e5c5c-610">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-610">Name</span></span>| <span data-ttu-id="e5c5c-611">種類</span><span class="sxs-lookup"><span data-stu-id="e5c5c-611">Type</span></span>| <span data-ttu-id="e5c5c-612">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e5c5c-613">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-613">String &#124; Object</span></span>| |<span data-ttu-id="e5c5c-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5c5c-616">**または**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-616">**OR**</span></span><br/><span data-ttu-id="e5c5c-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e5c5c-619">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-619">String</span></span> | <span data-ttu-id="e5c5c-620">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-620">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c5c-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="e5c5c-623">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-623">function</span></span> | <span data-ttu-id="e5c5c-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-624">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c5c-625">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5c5c-626">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-626">Requirements</span></span>

|<span data-ttu-id="e5c5c-627">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-627">Requirement</span></span>| <span data-ttu-id="e5c5c-628">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-629">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-630">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-630">1.0</span></span>|
|[<span data-ttu-id="e5c5c-631">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-632">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-633">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-634">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c5c-635">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-635">Examples</span></span>

<span data-ttu-id="e5c5c-636">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e5c5c-637">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e5c5c-638">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5c5c-639">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-639">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="e5c5c-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="e5c5c-641">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-642">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5c5c-643">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e5c5c-644">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-645">`displayReplyForm` に対する呼び出しに添付ファイルを含める機能は、要件セット 1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="e5c5c-646">添付ファイルのサポートは、要件セット 1.2 以降で `displayReplyForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-647">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-647">Parameters:</span></span>

|<span data-ttu-id="e5c5c-648">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-648">Name</span></span>| <span data-ttu-id="e5c5c-649">種類</span><span class="sxs-lookup"><span data-stu-id="e5c5c-649">Type</span></span>| <span data-ttu-id="e5c5c-650">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="e5c5c-651">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-651">String &#124; Object</span></span>| | <span data-ttu-id="e5c5c-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e5c5c-654">**または**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-654">**OR**</span></span><br/><span data-ttu-id="e5c5c-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="e5c5c-657">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-657">String</span></span> | <span data-ttu-id="e5c5c-658">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-658">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c5c-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="e5c5c-661">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-661">function</span></span> | <span data-ttu-id="e5c5c-662">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-662">&lt;optional&gt;</span></span> | <span data-ttu-id="e5c5c-663">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5c5c-664">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-664">Requirements</span></span>

|<span data-ttu-id="e5c5c-665">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-665">Requirement</span></span>| <span data-ttu-id="e5c5c-666">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-667">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-668">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-668">1.0</span></span>|
|[<span data-ttu-id="e5c5c-669">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-670">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-671">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-672">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e5c5c-673">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-673">Examples</span></span>

<span data-ttu-id="e5c5c-674">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e5c5c-675">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e5c5c-676">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e5c5c-677">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-677">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="e5c5c-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e5c5c-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="e5c5c-679">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-680">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-681">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-681">Requirements</span></span>

|<span data-ttu-id="e5c5c-682">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-682">Requirement</span></span>| <span data-ttu-id="e5c5c-683">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-684">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-685">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-685">1.0</span></span>|
|[<span data-ttu-id="e5c5c-686">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-687">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-688">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-689">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c5c-690">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-690">Returns:</span></span>

<span data-ttu-id="e5c5c-691">型:[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e5c5c-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e5c5c-692">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-692">Example</span></span>

<span data-ttu-id="e5c5c-693">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="e5c5c-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5c5c-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5c5c-695">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-696">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-697">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-697">Parameters:</span></span>

|<span data-ttu-id="e5c5c-698">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-698">Name</span></span>| <span data-ttu-id="e5c5c-699">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-699">Type</span></span>| <span data-ttu-id="e5c5c-700">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="e5c5c-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e5c5c-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="e5c5c-702">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c5c-703">Requirements</span><span class="sxs-lookup"><span data-stu-id="e5c5c-703">Requirements</span></span>

|<span data-ttu-id="e5c5c-704">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-704">Requirement</span></span>| <span data-ttu-id="e5c5c-705">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-706">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-707">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-707">1.0</span></span>|
|[<span data-ttu-id="e5c5c-708">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-709">制限あり</span><span class="sxs-lookup"><span data-stu-id="e5c5c-709">Restricted</span></span>|
|[<span data-ttu-id="e5c5c-710">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-711">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c5c-712">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-712">Returns:</span></span>

<span data-ttu-id="e5c5c-713">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e5c5c-714">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="e5c5c-715">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e5c5c-716">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="e5c5c-717">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-717">Value of `entityType`</span></span> | <span data-ttu-id="e5c5c-718">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-718">Type of objects in returned array</span></span> | <span data-ttu-id="e5c5c-719">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="e5c5c-720">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-720">String</span></span> | <span data-ttu-id="e5c5c-721">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="e5c5c-722">連絡先</span><span class="sxs-lookup"><span data-stu-id="e5c5c-722">Contact</span></span> | <span data-ttu-id="e5c5c-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="e5c5c-724">文字列</span><span class="sxs-lookup"><span data-stu-id="e5c5c-724">String</span></span> | <span data-ttu-id="e5c5c-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="e5c5c-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5c5c-726">MeetingSuggestion</span></span> | <span data-ttu-id="e5c5c-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="e5c5c-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e5c5c-728">PhoneNumber</span></span> | <span data-ttu-id="e5c5c-729">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="e5c5c-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e5c5c-730">TaskSuggestion</span></span> | <span data-ttu-id="e5c5c-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="e5c5c-732">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-732">String</span></span> | <span data-ttu-id="e5c5c-733">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="e5c5c-733">**Restricted**</span></span> |

<span data-ttu-id="e5c5c-734">型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5c5c-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="e5c5c-735">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-735">Example</span></span>

<span data-ttu-id="e5c5c-736">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="e5c5c-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e5c5c-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e5c5c-738">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-739">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5c5c-740">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-741">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-741">Parameters:</span></span>

|<span data-ttu-id="e5c5c-742">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-742">Name</span></span>| <span data-ttu-id="e5c5c-743">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-743">Type</span></span>| <span data-ttu-id="e5c5c-744">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e5c5c-745">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-745">String</span></span>|<span data-ttu-id="e5c5c-746">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c5c-747">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-747">Requirements</span></span>

|<span data-ttu-id="e5c5c-748">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-748">Requirement</span></span>| <span data-ttu-id="e5c5c-749">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-750">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-751">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-751">1.0</span></span>|
|[<span data-ttu-id="e5c5c-752">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-753">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-754">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-755">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c5c-756">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-756">Returns:</span></span>

<span data-ttu-id="e5c5c-p146">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="e5c5c-759">型:Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e5c5c-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="e5c5c-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e5c5c-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e5c5c-761">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-762">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5c5c-p147">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e5c5c-766">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e5c5c-767">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="e5c5c-p148">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e5c5c-770">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-770">Requirements</span></span>

|<span data-ttu-id="e5c5c-771">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-771">Requirement</span></span>| <span data-ttu-id="e5c5c-772">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-773">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-774">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-774">1.0</span></span>|
|[<span data-ttu-id="e5c5c-775">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-776">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-777">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-778">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c5c-779">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-779">Returns:</span></span>

<span data-ttu-id="e5c5c-p149">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e5c5c-782">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5c5c-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5c5c-783">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e5c5c-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5c5c-784">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-784">Example</span></span>

<span data-ttu-id="e5c5c-785">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="e5c5c-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e5c5c-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e5c5c-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e5c5c-787">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e5c5c-788">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e5c5c-789">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e5c5c-p150">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-792">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-792">Parameters:</span></span>

|<span data-ttu-id="e5c5c-793">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-793">Name</span></span>| <span data-ttu-id="e5c5c-794">種類</span><span class="sxs-lookup"><span data-stu-id="e5c5c-794">Type</span></span>| <span data-ttu-id="e5c5c-795">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="e5c5c-796">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-796">String</span></span>|<span data-ttu-id="e5c5c-797">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c5c-798">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-798">Requirements</span></span>

|<span data-ttu-id="e5c5c-799">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-799">Requirement</span></span>| <span data-ttu-id="e5c5c-800">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-801">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-802">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-802">1.0</span></span>|
|[<span data-ttu-id="e5c5c-803">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-804">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-805">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-806">読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e5c5c-807">戻り値:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-807">Returns:</span></span>

<span data-ttu-id="e5c5c-808">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="e5c5c-809">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e5c5c-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e5c5c-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e5c5c-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e5c5c-811">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e5c5c-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e5c5c-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e5c5c-813">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e5c5c-p151">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-817">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-817">Parameters:</span></span>

|<span data-ttu-id="e5c5c-818">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-818">Name</span></span>| <span data-ttu-id="e5c5c-819">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-819">Type</span></span>| <span data-ttu-id="e5c5c-820">属性</span><span class="sxs-lookup"><span data-stu-id="e5c5c-820">Attributes</span></span>| <span data-ttu-id="e5c5c-821">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e5c5c-822">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-822">function</span></span>||<span data-ttu-id="e5c5c-823">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e5c5c-824">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e5c5c-825">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="e5c5c-826">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-826">Object</span></span>| <span data-ttu-id="e5c5c-827">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-827">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-828">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="e5c5c-829">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e5c5c-830">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-830">Requirements</span></span>

|<span data-ttu-id="e5c5c-831">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-831">Requirement</span></span>| <span data-ttu-id="e5c5c-832">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-833">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-834">1.0</span><span class="sxs-lookup"><span data-stu-id="e5c5c-834">1.0</span></span>|
|[<span data-ttu-id="e5c5c-835">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-836">ReadItem</span></span>|
|[<span data-ttu-id="e5c5c-837">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-838">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="e5c5c-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-839">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-839">Example</span></span>

<span data-ttu-id="e5c5c-p154">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e5c5c-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e5c5c-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e5c5c-844">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e5c5c-p155">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e5c5c-849">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="e5c5c-849">Parameters:</span></span>

|<span data-ttu-id="e5c5c-850">名前</span><span class="sxs-lookup"><span data-stu-id="e5c5c-850">Name</span></span>| <span data-ttu-id="e5c5c-851">型</span><span class="sxs-lookup"><span data-stu-id="e5c5c-851">Type</span></span>| <span data-ttu-id="e5c5c-852">属性</span><span class="sxs-lookup"><span data-stu-id="e5c5c-852">Attributes</span></span>| <span data-ttu-id="e5c5c-853">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="e5c5c-854">String</span><span class="sxs-lookup"><span data-stu-id="e5c5c-854">String</span></span>||<span data-ttu-id="e5c5c-855">削除する添付ファイルの識別子。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="e5c5c-856">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="e5c5c-856">Object</span></span>| <span data-ttu-id="e5c5c-857">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-857">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-858">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="e5c5c-859">Object</span><span class="sxs-lookup"><span data-stu-id="e5c5c-859">Object</span></span>| <span data-ttu-id="e5c5c-860">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-860">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-861">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="e5c5c-862">function</span><span class="sxs-lookup"><span data-stu-id="e5c5c-862">function</span></span>| <span data-ttu-id="e5c5c-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e5c5c-863">&lt;optional&gt;</span></span>|<span data-ttu-id="e5c5c-864">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e5c5c-865">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e5c5c-866">エラー</span><span class="sxs-lookup"><span data-stu-id="e5c5c-866">Errors</span></span>

| <span data-ttu-id="e5c5c-867">エラー コード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-867">Error code</span></span> | <span data-ttu-id="e5c5c-868">説明</span><span class="sxs-lookup"><span data-stu-id="e5c5c-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="e5c5c-869">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e5c5c-870">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-870">Requirements</span></span>

|<span data-ttu-id="e5c5c-871">要件</span><span class="sxs-lookup"><span data-stu-id="e5c5c-871">Requirement</span></span>| <span data-ttu-id="e5c5c-872">値</span><span class="sxs-lookup"><span data-stu-id="e5c5c-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="e5c5c-873">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e5c5c-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e5c5c-874">1.1</span><span class="sxs-lookup"><span data-stu-id="e5c5c-874">1.1</span></span>|
|[<span data-ttu-id="e5c5c-875">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e5c5c-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e5c5c-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e5c5c-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="e5c5c-877">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e5c5c-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e5c5c-878">作成</span><span class="sxs-lookup"><span data-stu-id="e5c5c-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e5c5c-879">例</span><span class="sxs-lookup"><span data-stu-id="e5c5c-879">Example</span></span>

<span data-ttu-id="e5c5c-880">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="e5c5c-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
