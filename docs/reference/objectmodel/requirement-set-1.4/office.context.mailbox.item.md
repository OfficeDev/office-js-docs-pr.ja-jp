
# <a name="item"></a><span data-ttu-id="0f62e-101">item</span><span class="sxs-lookup"><span data-stu-id="0f62e-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0f62e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0f62e-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0f62e-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-105">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-105">Requirements</span></span>

|<span data-ttu-id="0f62e-106">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-106">Requirement</span></span>| <span data-ttu-id="0f62e-107">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-109">1.0</span></span>|
|[<span data-ttu-id="0f62e-110">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="0f62e-111">Restricted</span></span>|
|[<span data-ttu-id="0f62e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-113">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="0f62e-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="0f62e-114">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-114">Example</span></span>

<span data-ttu-id="0f62e-115">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0f62e-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="0f62e-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="0f62e-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0f62e-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="0f62e-p102">アイテムの添付ファイルの配列を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-120">セキュリティ上の問題がある可能性があるため、特定の種類のファイルは Outlook によってブロックされるので、返されません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0f62e-121">詳細については、「[Outlook でブロックされる添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f62e-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-122">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-122">Type:</span></span>

*   <span data-ttu-id="0f62e-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="0f62e-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-124">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-124">Requirements</span></span>

|<span data-ttu-id="0f62e-125">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-125">Requirement</span></span>| <span data-ttu-id="0f62e-126">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-127">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-128">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-128">1.0</span></span>|
|[<span data-ttu-id="0f62e-129">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-130">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-131">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-132">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-133">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-133">Example</span></span>

<span data-ttu-id="0f62e-134">次のコードでは、現在のアイテムのすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="0f62e-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="0f62e-136">メッセージの BCC (ブラインド カーボン コピー) 行を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0f62e-137">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-138">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-138">Type:</span></span>

*   [<span data-ttu-id="0f62e-139">Recipients</span><span class="sxs-lookup"><span data-stu-id="0f62e-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="0f62e-140">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-140">Requirements</span></span>

|<span data-ttu-id="0f62e-141">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-141">Requirement</span></span>| <span data-ttu-id="0f62e-142">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-143">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-144">1.1</span><span class="sxs-lookup"><span data-stu-id="0f62e-144">1.1</span></span>|
|[<span data-ttu-id="0f62e-145">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-146">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-147">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-148">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-149">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-149">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="0f62e-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="0f62e-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="0f62e-151">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-152">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-152">Type:</span></span>

*   [<span data-ttu-id="0f62e-153">Body</span><span class="sxs-lookup"><span data-stu-id="0f62e-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="0f62e-154">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-154">Requirements</span></span>

|<span data-ttu-id="0f62e-155">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-155">Requirement</span></span>| <span data-ttu-id="0f62e-156">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-157">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-158">1.1</span><span class="sxs-lookup"><span data-stu-id="0f62e-158">1.1</span></span>|
|[<span data-ttu-id="0f62e-159">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-160">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="0f62e-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="0f62e-164">メッセージの CC (カーボン コピー) の受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0f62e-165">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-166">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-166">Read mode</span></span>

<span data-ttu-id="0f62e-p106">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-169">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-169">Compose mode</span></span>

<span data-ttu-id="0f62e-170">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-171">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-171">Type:</span></span>

*   <span data-ttu-id="0f62e-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-173">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-173">Requirements</span></span>

|<span data-ttu-id="0f62e-174">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-174">Requirement</span></span>| <span data-ttu-id="0f62e-175">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-177">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-177">1.0</span></span>|
|[<span data-ttu-id="0f62e-178">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-179">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-182">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-182">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="0f62e-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="0f62e-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="0f62e-184">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0f62e-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0f62e-p108">新規作成フォームで新しいアイテムに対してこのプロパティに null を取得します。ユーザーが件名を設定し、アイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-189">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-189">Type:</span></span>

*   <span data-ttu-id="0f62e-190">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-191">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-191">Requirements</span></span>

|<span data-ttu-id="0f62e-192">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-192">Requirement</span></span>| <span data-ttu-id="0f62e-193">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-195">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-195">1.0</span></span>|
|[<span data-ttu-id="0f62e-196">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-197">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-199">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="0f62e-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="0f62e-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="0f62e-p109">アイテムが作成された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-203">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-203">Type:</span></span>

*   <span data-ttu-id="0f62e-204">日付</span><span class="sxs-lookup"><span data-stu-id="0f62e-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-205">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-205">Requirements</span></span>

|<span data-ttu-id="0f62e-206">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-206">Requirement</span></span>| <span data-ttu-id="0f62e-207">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-209">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-209">1.0</span></span>|
|[<span data-ttu-id="0f62e-210">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-211">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-213">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-214">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-214">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="0f62e-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="0f62e-215">dateTimeModified :Date</span></span>

<span data-ttu-id="0f62e-p110">アイテムが最後に変更された日時を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-218">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-219">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-219">Type:</span></span>

*   <span data-ttu-id="0f62e-220">日付</span><span class="sxs-lookup"><span data-stu-id="0f62e-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-221">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-221">Requirements</span></span>

|<span data-ttu-id="0f62e-222">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-222">Requirement</span></span>| <span data-ttu-id="0f62e-223">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-225">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-225">1.0</span></span>|
|[<span data-ttu-id="0f62e-226">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-227">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-229">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-230">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-230">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="0f62e-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="0f62e-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="0f62e-232">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0f62e-p111">`end` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-235">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-235">Read mode</span></span>

<span data-ttu-id="0f62e-236">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-237">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-237">Compose mode</span></span>

<span data-ttu-id="0f62e-238">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0f62e-239">[`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-240">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-240">Type:</span></span>

*   <span data-ttu-id="0f62e-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="0f62e-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-242">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-242">Requirements</span></span>

|<span data-ttu-id="0f62e-243">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-243">Requirement</span></span>| <span data-ttu-id="0f62e-244">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-246">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-246">1.0</span></span>|
|[<span data-ttu-id="0f62e-247">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-248">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-251">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-251">Example</span></span>

<span data-ttu-id="0f62e-252">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="0f62e-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0f62e-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="0f62e-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="0f62e-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-258">`from` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-259">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-259">Type:</span></span>

*   [<span data-ttu-id="0f62e-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f62e-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0f62e-261">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-261">Requirements</span></span>

|<span data-ttu-id="0f62e-262">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-262">Requirement</span></span>| <span data-ttu-id="0f62e-263">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-265">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-265">1.0</span></span>|
|[<span data-ttu-id="0f62e-266">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-267">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-269">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="0f62e-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="0f62e-270">internetMessageId :String</span></span>

<span data-ttu-id="0f62e-p114">電子メール メッセージのインターネット メッセージ ID を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-273">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-273">Type:</span></span>

*   <span data-ttu-id="0f62e-274">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-275">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-275">Requirements</span></span>

|<span data-ttu-id="0f62e-276">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-276">Requirement</span></span>| <span data-ttu-id="0f62e-277">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-279">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-279">1.0</span></span>|
|[<span data-ttu-id="0f62e-280">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-281">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-283">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-284">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-284">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="0f62e-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="0f62e-285">itemClass :String</span></span>

<span data-ttu-id="0f62e-p115">選択されたアイテムの Exchange Web サービスのアイテム クラスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0f62e-p116">`itemClass` プロパティには、選択したアイテムのメッセージ クラスを指定します。次に、メッセージまたは予定アイテムの既定のメッセージ クラスを示します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="0f62e-290">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-290">Type</span></span> | <span data-ttu-id="0f62e-291">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-291">Description</span></span> | <span data-ttu-id="0f62e-292">アイテム クラス</span><span class="sxs-lookup"><span data-stu-id="0f62e-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="0f62e-293">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="0f62e-293">Appointment items</span></span> | <span data-ttu-id="0f62e-294">アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="0f62e-295">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="0f62e-295">Message items</span></span> | <span data-ttu-id="0f62e-296">これには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージ、および基本メッセージ クラスとして `IPM.Schedule.Meeting` を使用する会議出席依頼、返信、または取り消しが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="0f62e-297">既定のメッセージ クラスを拡張したカスタム メッセージ クラス (たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など) を作成できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-298">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-298">Type:</span></span>

*   <span data-ttu-id="0f62e-299">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-300">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-300">Requirements</span></span>

|<span data-ttu-id="0f62e-301">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-301">Requirement</span></span>| <span data-ttu-id="0f62e-302">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-304">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-304">1.0</span></span>|
|[<span data-ttu-id="0f62e-305">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-306">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-308">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-309">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-309">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0f62e-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="0f62e-310">(nullable) itemId :String</span></span>

<span data-ttu-id="0f62e-p117">現在のアイテムの Exchange Web サービスのアイテム識別子を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-313">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0f62e-314">`itemId` プロパティは、Outlook Entry ID または Outlook REST API で使用される ID と同一ではありません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0f62e-315">この値を使用して REST API を呼び出す前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0f62e-316">詳細については、「[Outlook アドインから Outlook REST API を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0f62e-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0f62e-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-319">種類:</span><span class="sxs-lookup"><span data-stu-id="0f62e-319">Type:</span></span>

*   <span data-ttu-id="0f62e-320">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-321">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-321">Requirements</span></span>

|<span data-ttu-id="0f62e-322">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-322">Requirement</span></span>| <span data-ttu-id="0f62e-323">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-325">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-325">1.0</span></span>|
|[<span data-ttu-id="0f62e-326">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-327">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-329">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-330">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-330">Example</span></span>

<span data-ttu-id="0f62e-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="0f62e-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="0f62e-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="0f62e-334">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0f62e-335">`itemType` プロパティは、`ItemType` 列挙値の 1 つを返します。これは `item` オブジェクト インスタンスがメッセージと予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-336">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-336">Type:</span></span>

*   [<span data-ttu-id="0f62e-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0f62e-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="0f62e-338">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-338">Requirements</span></span>

|<span data-ttu-id="0f62e-339">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-339">Requirement</span></span>| <span data-ttu-id="0f62e-340">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-342">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-342">1.0</span></span>|
|[<span data-ttu-id="0f62e-343">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-344">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-346">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-347">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-347">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="0f62e-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="0f62e-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="0f62e-349">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-350">Read mode</span></span>

<span data-ttu-id="0f62e-351">`location` プロパティは、予定の場所を格納した文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-352">Compose mode</span></span>

<span data-ttu-id="0f62e-353">`location` プロパティは予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-354">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-354">Type:</span></span>

*   <span data-ttu-id="0f62e-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="0f62e-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-356">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-356">Requirements</span></span>

|<span data-ttu-id="0f62e-357">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-357">Requirement</span></span>| <span data-ttu-id="0f62e-358">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-360">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-360">1.0</span></span>|
|[<span data-ttu-id="0f62e-361">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-362">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-364">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-365">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-365">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0f62e-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="0f62e-366">normalizedSubject :String</span></span>

<span data-ttu-id="0f62e-p121">すべてのプレフィックス (`RE:` や `FWD:` など) が削除されたアイテムの件名を取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0f62e-p122">normalizedSubject プロパティは、アイテムの件名に電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたものを取得します。これらのプレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-371">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-371">Type:</span></span>

*   <span data-ttu-id="0f62e-372">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-373">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-373">Requirements</span></span>

|<span data-ttu-id="0f62e-374">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-374">Requirement</span></span>| <span data-ttu-id="0f62e-375">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-376">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-377">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-377">1.0</span></span>|
|[<span data-ttu-id="0f62e-378">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-379">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-381">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-382">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-382">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="0f62e-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="0f62e-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="0f62e-384">アイテムの通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-385">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-385">Type:</span></span>

*   [<span data-ttu-id="0f62e-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0f62e-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="0f62e-387">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-387">Requirements</span></span>

|<span data-ttu-id="0f62e-388">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-388">Requirement</span></span>| <span data-ttu-id="0f62e-389">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-391">1.3</span><span class="sxs-lookup"><span data-stu-id="0f62e-391">1.3</span></span>|
|[<span data-ttu-id="0f62e-392">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-393">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-395">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="0f62e-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="0f62e-397">イベントの任意出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0f62e-398">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-399">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-399">Read mode</span></span>

<span data-ttu-id="0f62e-400">`optionalAttendees` プロパティは、会議への各任意出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-401">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-401">Compose mode</span></span>

<span data-ttu-id="0f62e-402">`optionalAttendees` プロパティは会議への任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-403">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-403">Type:</span></span>

*   <span data-ttu-id="0f62e-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-405">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-405">Requirements</span></span>

|<span data-ttu-id="0f62e-406">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-406">Requirement</span></span>| <span data-ttu-id="0f62e-407">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-408">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-409">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-409">1.0</span></span>|
|[<span data-ttu-id="0f62e-410">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-411">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-412">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-413">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-414">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-414">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="0f62e-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0f62e-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="0f62e-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-418">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-418">Type:</span></span>

*   [<span data-ttu-id="0f62e-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f62e-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0f62e-420">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-420">Requirements</span></span>

|<span data-ttu-id="0f62e-421">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-421">Requirement</span></span>| <span data-ttu-id="0f62e-422">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-424">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-424">1.0</span></span>|
|[<span data-ttu-id="0f62e-425">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-426">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-428">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-429">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-429">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="0f62e-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="0f62e-431">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0f62e-432">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-433">Read mode</span></span>

<span data-ttu-id="0f62e-434">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを格納した配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-435">Compose mode</span></span>

<span data-ttu-id="0f62e-436">`requiredAttendees` プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-437">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-437">Type:</span></span>

*   <span data-ttu-id="0f62e-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-439">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-439">Requirements</span></span>

|<span data-ttu-id="0f62e-440">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-440">Requirement</span></span>| <span data-ttu-id="0f62e-441">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-443">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-443">1.0</span></span>|
|[<span data-ttu-id="0f62e-444">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-445">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-447">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-448">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-448">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="0f62e-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="0f62e-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="0f62e-p126">電子メール メッセージの送信者の電子メール アドレスを取得します。閲覧モードのみ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0f62e-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。代理人から送信された場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-454">`sender` プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType` プロパティは `undefined` です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-455">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-455">Type:</span></span>

*   [<span data-ttu-id="0f62e-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0f62e-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="0f62e-457">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-457">Requirements</span></span>

|<span data-ttu-id="0f62e-458">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-458">Requirement</span></span>| <span data-ttu-id="0f62e-459">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-461">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-461">1.0</span></span>|
|[<span data-ttu-id="0f62e-462">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-463">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-465">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-466">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-466">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="0f62e-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="0f62e-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="0f62e-468">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0f62e-p128">`start` プロパティは、世界協定時刻 (UTC) 形式の日時値として表されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-471">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-471">Read mode</span></span>

<span data-ttu-id="0f62e-472">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-473">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-473">Compose mode</span></span>

<span data-ttu-id="0f62e-474">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0f62e-475">[`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-476">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-476">Type:</span></span>

*   <span data-ttu-id="0f62e-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="0f62e-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-478">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-478">Requirements</span></span>

|<span data-ttu-id="0f62e-479">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-479">Requirement</span></span>| <span data-ttu-id="0f62e-480">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-482">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-482">1.0</span></span>|
|[<span data-ttu-id="0f62e-483">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-484">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-486">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-487">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-487">Example</span></span>

<span data-ttu-id="0f62e-488">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="0f62e-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0f62e-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="0f62e-490">アイテムの件名フィールドに示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0f62e-491">`subject` プロパティは、電子メール サーバーによって送信されたアイテムの件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-492">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-492">Read mode</span></span>

<span data-ttu-id="0f62e-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:` や `FW:` のような先頭部分のすべてのプレフィックスを除去した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-495">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-495">Compose mode</span></span>

<span data-ttu-id="0f62e-496">`subject` プロパティは件名を取得および設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0f62e-497">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-497">Type:</span></span>

*   <span data-ttu-id="0f62e-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="0f62e-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-499">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-499">Requirements</span></span>

|<span data-ttu-id="0f62e-500">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-500">Requirement</span></span>| <span data-ttu-id="0f62e-501">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-503">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-503">1.0</span></span>|
|[<span data-ttu-id="0f62e-504">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-505">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-507">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="0f62e-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="0f62e-509">メッセージの **To** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0f62e-510">オブジェクトの種類とアクセスのレベルは、現在のアイテムのモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0f62e-511">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-511">Read mode</span></span>

<span data-ttu-id="0f62e-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="0f62e-514">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="0f62e-514">Compose mode</span></span>

<span data-ttu-id="0f62e-515">`to` プロパティは、メッセージの **To** 行の受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="0f62e-516">型:</span><span class="sxs-lookup"><span data-stu-id="0f62e-516">Type:</span></span>

*   <span data-ttu-id="0f62e-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="0f62e-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-518">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-518">Requirements</span></span>

|<span data-ttu-id="0f62e-519">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-519">Requirement</span></span>| <span data-ttu-id="0f62e-520">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-522">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-522">1.0</span></span>|
|[<span data-ttu-id="0f62e-523">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-524">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-526">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-527">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-527">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="0f62e-528">メソッド</span><span class="sxs-lookup"><span data-stu-id="0f62e-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0f62e-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f62e-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0f62e-530">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0f62e-531">`addFileAttachmentAsync` メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内のアイテムに添付します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0f62e-532">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-533">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-533">Parameters:</span></span>

|<span data-ttu-id="0f62e-534">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-534">Name</span></span>| <span data-ttu-id="0f62e-535">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-535">Type</span></span>| <span data-ttu-id="0f62e-536">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-536">Attributes</span></span>| <span data-ttu-id="0f62e-537">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="0f62e-538">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-538">String</span></span>||<span data-ttu-id="0f62e-p132">メッセージまたは予定に添付するファイルの場所を示す URI。最大長は 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0f62e-541">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-541">String</span></span>||<span data-ttu-id="0f62e-p133">添付ファイルのアップロード時に表示される添付ファイルの名前。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0f62e-544">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-544">Object</span></span>| <span data-ttu-id="0f62e-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-545">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-546">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-547">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-547">Object</span></span>| <span data-ttu-id="0f62e-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-548">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-549">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f62e-550">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-550">function</span></span>| <span data-ttu-id="0f62e-551">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-551">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-552">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f62e-553">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0f62e-554">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f62e-555">エラー</span><span class="sxs-lookup"><span data-stu-id="0f62e-555">Errors</span></span>

| <span data-ttu-id="0f62e-556">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f62e-556">Error code</span></span> | <span data-ttu-id="0f62e-557">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="0f62e-558">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="0f62e-559">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0f62e-560">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-561">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-561">Requirements</span></span>

|<span data-ttu-id="0f62e-562">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-562">Requirement</span></span>| <span data-ttu-id="0f62e-563">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-565">1.1</span><span class="sxs-lookup"><span data-stu-id="0f62e-565">1.1</span></span>|
|[<span data-ttu-id="0f62e-566">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-569">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-570">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0f62e-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f62e-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0f62e-572">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0f62e-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つアイテムを新規作成フォーム内のアイテムに添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメーターがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、またはアイテムの添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメーターを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0f62e-576">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0f62e-577">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドはアイテムを、編集中のアイテム以外のアイテムに添付できますが、これはサポートされておらず、お勧めできません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-578">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-578">Parameters:</span></span>

|<span data-ttu-id="0f62e-579">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-579">Name</span></span>| <span data-ttu-id="0f62e-580">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-580">Type</span></span>| <span data-ttu-id="0f62e-581">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-581">Attributes</span></span>| <span data-ttu-id="0f62e-582">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="0f62e-583">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-583">String</span></span>||<span data-ttu-id="0f62e-p135">添付するアイテムの Exchange 識別子。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0f62e-586">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-586">String</span></span>||<span data-ttu-id="0f62e-p136">添付するアイテムの件名。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0f62e-589">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-589">Object</span></span>| <span data-ttu-id="0f62e-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-590">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-591">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-592">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-592">Object</span></span>| <span data-ttu-id="0f62e-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-593">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-594">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f62e-595">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-595">function</span></span>| <span data-ttu-id="0f62e-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-596">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-597">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f62e-598">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0f62e-599">添付ファイルの追加に失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f62e-600">エラー</span><span class="sxs-lookup"><span data-stu-id="0f62e-600">Errors</span></span>

| <span data-ttu-id="0f62e-601">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f62e-601">Error code</span></span> | <span data-ttu-id="0f62e-602">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0f62e-603">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-604">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-604">Requirements</span></span>

|<span data-ttu-id="0f62e-605">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-605">Requirement</span></span>| <span data-ttu-id="0f62e-606">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-608">1.1</span><span class="sxs-lookup"><span data-stu-id="0f62e-608">1.1</span></span>|
|[<span data-ttu-id="0f62e-609">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-612">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-613">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-613">Example</span></span>

<span data-ttu-id="0f62e-614">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="0f62e-615">close()</span><span class="sxs-lookup"><span data-stu-id="0f62e-615">close()</span></span>

<span data-ttu-id="0f62e-616">作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0f62e-p137">`close` メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-619">Outlook on the web では、アイテムが予定であり、`saveAsync` を利用して以前に保存されている場合、アイテムが最後に保存された後に変更が行われていなくても、保存、破棄、またはキャンセルするようにユーザーは求められます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0f62e-620">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close` メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-621">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-621">Requirements</span></span>

|<span data-ttu-id="0f62e-622">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-622">Requirement</span></span>| <span data-ttu-id="0f62e-623">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-624">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-625">1.3</span><span class="sxs-lookup"><span data-stu-id="0f62e-625">1.3</span></span>|
|[<span data-ttu-id="0f62e-626">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-627">制限あり</span><span class="sxs-lookup"><span data-stu-id="0f62e-627">Restricted</span></span>|
|[<span data-ttu-id="0f62e-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-629">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="0f62e-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="0f62e-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="0f62e-631">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-632">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f62e-633">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0f62e-634">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="0f62e-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0f62e-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-638">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-638">Parameters:</span></span>

|<span data-ttu-id="0f62e-639">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-639">Name</span></span>| <span data-ttu-id="0f62e-640">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-640">Type</span></span>| <span data-ttu-id="0f62e-641">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0f62e-642">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-642">String &#124; Object</span></span>| |<span data-ttu-id="0f62e-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0f62e-645">**または**</span><span class="sxs-lookup"><span data-stu-id="0f62e-645">**OR**</span></span><br/><span data-ttu-id="0f62e-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0f62e-648">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-648">String</span></span> | <span data-ttu-id="0f62e-649">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-649">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0f62e-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0f62e-653">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-653">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-654">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0f62e-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0f62e-655">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-655">String</span></span> | | <span data-ttu-id="0f62e-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0f62e-658">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-658">String</span></span> | | <span data-ttu-id="0f62e-659">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0f62e-660">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-660">String</span></span> | | <span data-ttu-id="0f62e-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0f62e-663">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-663">String</span></span> | | <span data-ttu-id="0f62e-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0f62e-667">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-667">function</span></span> | <span data-ttu-id="0f62e-668">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-668">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-669">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-670">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-670">Requirements</span></span>

|<span data-ttu-id="0f62e-671">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-671">Requirement</span></span>| <span data-ttu-id="0f62e-672">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-673">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-674">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-674">1.0</span></span>|
|[<span data-ttu-id="0f62e-675">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-676">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-678">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0f62e-679">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-679">Examples</span></span>

<span data-ttu-id="0f62e-680">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0f62e-681">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-681">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0f62e-682">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-682">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0f62e-683">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0f62e-684">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0f62e-685">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="0f62e-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="0f62e-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="0f62e-687">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-688">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f62e-689">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0f62e-690">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="0f62e-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0f62e-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-694">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-694">Parameters:</span></span>

|<span data-ttu-id="0f62e-695">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-695">Name</span></span>| <span data-ttu-id="0f62e-696">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-696">Type</span></span>| <span data-ttu-id="0f62e-697">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0f62e-698">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-698">String &#124; Object</span></span>| | <span data-ttu-id="0f62e-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0f62e-701">**または**</span><span class="sxs-lookup"><span data-stu-id="0f62e-701">**OR**</span></span><br/><span data-ttu-id="0f62e-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクト。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0f62e-704">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-704">String</span></span> | <span data-ttu-id="0f62e-705">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-705">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="0f62e-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="0f62e-709">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-709">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-710">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="0f62e-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="0f62e-711">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-711">String</span></span> | | <span data-ttu-id="0f62e-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="0f62e-714">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-714">String</span></span> | | <span data-ttu-id="0f62e-715">添付ファイル名を含む文字列。最大の長さは 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="0f62e-716">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-716">String</span></span> | | <span data-ttu-id="0f62e-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URI。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="0f62e-719">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-719">String</span></span> | | <span data-ttu-id="0f62e-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムの ID。最大の長さが 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="0f62e-723">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-723">function</span></span> | <span data-ttu-id="0f62e-724">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-724">&lt;optional&gt;</span></span> | <span data-ttu-id="0f62e-725">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-726">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-726">Requirements</span></span>

|<span data-ttu-id="0f62e-727">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-727">Requirement</span></span>| <span data-ttu-id="0f62e-728">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-729">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-730">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-730">1.0</span></span>|
|[<span data-ttu-id="0f62e-731">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-732">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-734">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0f62e-735">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-735">Examples</span></span>

<span data-ttu-id="0f62e-736">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0f62e-737">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-737">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0f62e-738">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-738">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0f62e-739">本文とファイルの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0f62e-740">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0f62e-741">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="0f62e-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="0f62e-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="0f62e-743">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-744">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-745">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-745">Requirements</span></span>

|<span data-ttu-id="0f62e-746">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-746">Requirement</span></span>| <span data-ttu-id="0f62e-747">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-749">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-749">1.0</span></span>|
|[<span data-ttu-id="0f62e-750">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-751">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-753">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-754">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-754">Returns:</span></span>

<span data-ttu-id="0f62e-755">型:[Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="0f62e-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="0f62e-756">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-756">Example</span></span>

<span data-ttu-id="0f62e-757">次の例は、現在のアイテムの本文にある連絡先エンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="0f62e-757">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="0f62e-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0f62e-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0f62e-759">選択したアイテムの本文内で検出された指定のエンティティ型のすべてのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-760">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-761">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-761">Parameters:</span></span>

|<span data-ttu-id="0f62e-762">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-762">Name</span></span>| <span data-ttu-id="0f62e-763">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-763">Type</span></span>| <span data-ttu-id="0f62e-764">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="0f62e-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0f62e-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="0f62e-766">EntityType 列挙値の 1 つ。</span><span class="sxs-lookup"><span data-stu-id="0f62e-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-767">Requirements</span><span class="sxs-lookup"><span data-stu-id="0f62e-767">Requirements</span></span>

|<span data-ttu-id="0f62e-768">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-768">Requirement</span></span>| <span data-ttu-id="0f62e-769">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-770">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-771">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-771">1.0</span></span>|
|[<span data-ttu-id="0f62e-772">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-773">制限あり</span><span class="sxs-lookup"><span data-stu-id="0f62e-773">Restricted</span></span>|
|[<span data-ttu-id="0f62e-774">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-775">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-776">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-776">Returns:</span></span>

<span data-ttu-id="0f62e-777">`entityType` に渡された値が `EntityType` 列挙型の有効なメンバーでない場合、メソッドは null を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0f62e-778">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="0f62e-779">それ以外の場合は、返される配列内のオブジェクトの型は、`entityType` パラメーター内の要求されたエンティティの型によって異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0f62e-780">このメソッドを使用する最小限のアクセス許可レベルは **Restricted** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="0f62e-781">`entityType` の値</span><span class="sxs-lookup"><span data-stu-id="0f62e-781">Value of `entityType`</span></span> | <span data-ttu-id="0f62e-782">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="0f62e-782">Type of objects in returned array</span></span> | <span data-ttu-id="0f62e-783">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="0f62e-784">文字列</span><span class="sxs-lookup"><span data-stu-id="0f62e-784">String</span></span> | <span data-ttu-id="0f62e-785">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f62e-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="0f62e-786">連絡先</span><span class="sxs-lookup"><span data-stu-id="0f62e-786">Contact</span></span> | <span data-ttu-id="0f62e-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f62e-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="0f62e-788">文字列</span><span class="sxs-lookup"><span data-stu-id="0f62e-788">String</span></span> | <span data-ttu-id="0f62e-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f62e-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="0f62e-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0f62e-790">MeetingSuggestion</span></span> | <span data-ttu-id="0f62e-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f62e-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="0f62e-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0f62e-792">PhoneNumber</span></span> | <span data-ttu-id="0f62e-793">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f62e-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="0f62e-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0f62e-794">TaskSuggestion</span></span> | <span data-ttu-id="0f62e-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0f62e-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="0f62e-796">文字列</span><span class="sxs-lookup"><span data-stu-id="0f62e-796">String</span></span> | <span data-ttu-id="0f62e-797">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="0f62e-797">**Restricted**</span></span> |

<span data-ttu-id="0f62e-798">型:Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0f62e-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="0f62e-799">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-799">Example</span></span>

<span data-ttu-id="0f62e-800">次の例は、現在のアイテムの本文にある郵送先住所を表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="0f62e-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="0f62e-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="0f62e-802">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-803">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f62e-804">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-805">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-805">Parameters:</span></span>

|<span data-ttu-id="0f62e-806">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-806">Name</span></span>| <span data-ttu-id="0f62e-807">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-807">Type</span></span>| <span data-ttu-id="0f62e-808">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0f62e-809">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-809">String</span></span>|<span data-ttu-id="0f62e-810">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="0f62e-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-811">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-811">Requirements</span></span>

|<span data-ttu-id="0f62e-812">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-812">Requirement</span></span>| <span data-ttu-id="0f62e-813">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-814">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-815">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-815">1.0</span></span>|
|[<span data-ttu-id="0f62e-816">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-817">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-819">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-820">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-820">Returns:</span></span>

<span data-ttu-id="0f62e-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0f62e-823">型:Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="0f62e-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="0f62e-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0f62e-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0f62e-825">選択したアイテム内の、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-826">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f62e-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内の、各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定されたアイテムのプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0f62e-830">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="0f62e-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0f62e-831">`getRegExMatches` から返されるオブジェクトに `fruits` および `veggies` という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0f62e-p155">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。この場合、代わりに [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f62e-835">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-835">Requirements</span></span>

|<span data-ttu-id="0f62e-836">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-836">Requirement</span></span>| <span data-ttu-id="0f62e-837">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-838">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-839">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-839">1.0</span></span>|
|[<span data-ttu-id="0f62e-840">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-841">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-842">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-843">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-844">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-844">Returns:</span></span>

<span data-ttu-id="0f62e-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクト。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="0f62e-847">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f62e-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f62e-848">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f62e-849">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-849">Example</span></span>

<span data-ttu-id="0f62e-850">次の例は、マニフェストで指定された正規表現 <rule> の要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="0f62e-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0f62e-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="0f62e-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0f62e-852">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-853">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="0f62e-854">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0f62e-p157">アイテムの body プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-857">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-857">Parameters:</span></span>

|<span data-ttu-id="0f62e-858">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-858">Name</span></span>| <span data-ttu-id="0f62e-859">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-859">Type</span></span>| <span data-ttu-id="0f62e-860">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0f62e-861">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-861">String</span></span>|<span data-ttu-id="0f62e-862">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前。</span><span class="sxs-lookup"><span data-stu-id="0f62e-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-863">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-863">Requirements</span></span>

|<span data-ttu-id="0f62e-864">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-864">Requirement</span></span>| <span data-ttu-id="0f62e-865">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-867">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-867">1.0</span></span>|
|[<span data-ttu-id="0f62e-868">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-869">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-871">閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-872">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-872">Returns:</span></span>

<span data-ttu-id="0f62e-873">マニフェスト XML ファイルで定義された正規表現に一致する文字列が格納された配列。</span><span class="sxs-lookup"><span data-stu-id="0f62e-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="0f62e-874">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f62e-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f62e-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0f62e-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f62e-876">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0f62e-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0f62e-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0f62e-878">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0f62e-p158">選択したデータがなく、カーソルが本文または件名にある場合、選択したデータに対して null が返されます。本文または件名以外のフィールドが選択されている場合、`InvalidSelection` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-881">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-881">Parameters:</span></span>

|<span data-ttu-id="0f62e-882">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-882">Name</span></span>| <span data-ttu-id="0f62e-883">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-883">Type</span></span>| <span data-ttu-id="0f62e-884">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-884">Attributes</span></span>| <span data-ttu-id="0f62e-885">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="0f62e-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f62e-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0f62e-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="0f62e-890">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-890">Object</span></span>| <span data-ttu-id="0f62e-891">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-891">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-892">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-893">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-893">Object</span></span>| <span data-ttu-id="0f62e-894">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-894">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-895">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f62e-896">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-896">function</span></span>||<span data-ttu-id="0f62e-897">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0f62e-898">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data` を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0f62e-899">選択のソース プロパティにアクセスするには、`asyncResult.value.sourceProperty` を呼び出します。これは `body` または `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-900">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-900">Requirements</span></span>

|<span data-ttu-id="0f62e-901">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-901">Requirement</span></span>| <span data-ttu-id="0f62e-902">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-903">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-904">1.2</span><span class="sxs-lookup"><span data-stu-id="0f62e-904">1.2</span></span>|
|[<span data-ttu-id="0f62e-905">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-908">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0f62e-909">戻り値:</span><span class="sxs-lookup"><span data-stu-id="0f62e-909">Returns:</span></span>

<span data-ttu-id="0f62e-910">選択されたデータ (`coercionType` で決定された形式の文字列)。</span><span class="sxs-lookup"><span data-stu-id="0f62e-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="0f62e-911">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="0f62e-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="0f62e-912">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="0f62e-913">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0f62e-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0f62e-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0f62e-915">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0f62e-p161">カスタム プロパティは、アプリケーションごと、アイテムごとのキーと値のペアとして格納されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在のアイテムおよび現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、アイテム上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-919">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-919">Parameters:</span></span>

|<span data-ttu-id="0f62e-920">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-920">Name</span></span>| <span data-ttu-id="0f62e-921">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-921">Type</span></span>| <span data-ttu-id="0f62e-922">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-922">Attributes</span></span>| <span data-ttu-id="0f62e-923">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0f62e-924">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-924">function</span></span>||<span data-ttu-id="0f62e-925">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0f62e-926">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0f62e-927">このオブジェクトは、アイテムからカスタム プロパティを取得、設定、および削除し、サーバーに設定し直すカスタム プロパティへの変更を保存するために使用できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="0f62e-928">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-928">Object</span></span>| <span data-ttu-id="0f62e-929">&lt;省略可能&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-929">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-930">開発者は、コールバック関数でアクセスする任意のオブジェクトを指定できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="0f62e-931">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-932">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-932">Requirements</span></span>

|<span data-ttu-id="0f62e-933">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-933">Requirement</span></span>| <span data-ttu-id="0f62e-934">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-936">1.0</span><span class="sxs-lookup"><span data-stu-id="0f62e-936">1.0</span></span>|
|[<span data-ttu-id="0f62e-937">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-938">ReadItem</span></span>|
|[<span data-ttu-id="0f62e-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-940">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0f62e-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-941">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-941">Example</span></span>

<span data-ttu-id="0f62e-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0f62e-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0f62e-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0f62e-946">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0f62e-p165">`removeAttachmentAsync` メソッドは、指定した識別子の添付ファイルをアイテムから削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web App とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別ウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-951">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-951">Parameters:</span></span>

|<span data-ttu-id="0f62e-952">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-952">Name</span></span>| <span data-ttu-id="0f62e-953">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-953">Type</span></span>| <span data-ttu-id="0f62e-954">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-954">Attributes</span></span>| <span data-ttu-id="0f62e-955">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="0f62e-956">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-956">String</span></span>||<span data-ttu-id="0f62e-p166">削除する添付ファイルの識別子。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="0f62e-959">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-959">Object</span></span>| <span data-ttu-id="0f62e-960">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-960">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-961">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-962">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-962">Object</span></span>| <span data-ttu-id="0f62e-963">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-963">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-964">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0f62e-965">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-965">function</span></span>| <span data-ttu-id="0f62e-966">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-966">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-967">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0f62e-968">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0f62e-969">エラー</span><span class="sxs-lookup"><span data-stu-id="0f62e-969">Errors</span></span>

| <span data-ttu-id="0f62e-970">エラー コード</span><span class="sxs-lookup"><span data-stu-id="0f62e-970">Error code</span></span> | <span data-ttu-id="0f62e-971">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="0f62e-972">添付ファイル識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-973">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-973">Requirements</span></span>

|<span data-ttu-id="0f62e-974">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-974">Requirement</span></span>| <span data-ttu-id="0f62e-975">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-976">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-977">1.1</span><span class="sxs-lookup"><span data-stu-id="0f62e-977">1.1</span></span>|
|[<span data-ttu-id="0f62e-978">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-980">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-981">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-982">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-982">Example</span></span>

<span data-ttu-id="0f62e-983">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="0f62e-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0f62e-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="0f62e-985">項目を非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="0f62e-p167">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッドを使用してアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-989">EWS または REST API で使用するための `itemId` を取得するために、アドインが新規作成モードのアイテムで `saveAsync` を呼び出す場合、Outlook がキャッシュ モードになっていると、アイテムが実際にサーバーに同期されるまでに時間がかかる可能性があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="0f62e-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="0f62e-990">アイテムが同期されるまで、`itemId` を使用するとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0f62e-p169">予定はドラフト状態にはならないため、作成モードで予定に `saveAsync` が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0f62e-994">次のクライアントの場合、新規作成モードで予約の `saveAsync` に対して動作が異なります。</span><span class="sxs-lookup"><span data-stu-id="0f62e-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0f62e-995">Mac Outlook では、新規作成モードの会議で `saveAsync` をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="0f62e-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="0f62e-996">Mac Outlook では、会議で `saveAsync` を呼び出すとエラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="0f62e-997">Outlook on the web の場合、新規作成モードのとき、予約で `saveAsync` が呼び出されると、招待状または更新が常に送信されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-998">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-998">Parameters:</span></span>

|<span data-ttu-id="0f62e-999">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-999">Name</span></span>| <span data-ttu-id="0f62e-1000">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-1000">Type</span></span>| <span data-ttu-id="0f62e-1001">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-1001">Attributes</span></span>| <span data-ttu-id="0f62e-1002">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="0f62e-1003">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-1003">Object</span></span>| <span data-ttu-id="0f62e-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-1005">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-1006">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-1006">Object</span></span>| <span data-ttu-id="0f62e-1007">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-1008">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="0f62e-1009">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-1009">function</span></span>||<span data-ttu-id="0f62e-1010">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0f62e-1011">成功すると、アイテム識別子が `asyncResult.value` プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f62e-1012">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-1012">Requirements</span></span>

|<span data-ttu-id="0f62e-1013">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-1013">Requirement</span></span>| <span data-ttu-id="0f62e-1014">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-1015">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="0f62e-1016">1.3</span></span>|
|[<span data-ttu-id="0f62e-1017">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-1019">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-1020">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0f62e-1021">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="0f62e-p171">次の例は、コールバック関数に渡される `result` パラメーターの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0f62e-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0f62e-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0f62e-1025">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0f62e-p172">`setSelectedDataAsync` メソッドは、指定された文字列をアイテムのサブジェクトまたは本文のカーソル位置に挿入します。または、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0f62e-1029">パラメーター:</span><span class="sxs-lookup"><span data-stu-id="0f62e-1029">Parameters:</span></span>

|<span data-ttu-id="0f62e-1030">名前</span><span class="sxs-lookup"><span data-stu-id="0f62e-1030">Name</span></span>| <span data-ttu-id="0f62e-1031">型</span><span class="sxs-lookup"><span data-stu-id="0f62e-1031">Type</span></span>| <span data-ttu-id="0f62e-1032">属性</span><span class="sxs-lookup"><span data-stu-id="0f62e-1032">Attributes</span></span>| <span data-ttu-id="0f62e-1033">説明</span><span class="sxs-lookup"><span data-stu-id="0f62e-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="0f62e-1034">String</span><span class="sxs-lookup"><span data-stu-id="0f62e-1034">String</span></span>||<span data-ttu-id="0f62e-p173">挿入されるデータ。データの最大の長さは 1,000,000 文字です。1,000,000 文字を超えるデータが渡されると、`ArgumentOutOfRange` 例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="0f62e-1038">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-1038">Object</span></span>| <span data-ttu-id="0f62e-1039">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-1040">次のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0f62e-1041">Object</span><span class="sxs-lookup"><span data-stu-id="0f62e-1041">Object</span></span>| <span data-ttu-id="0f62e-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-1043">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="0f62e-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f62e-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="0f62e-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0f62e-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="0f62e-p174">`text` の場合、Office Web Apps と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0f62e-p175">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0f62e-1050">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="0f62e-1051">function</span><span class="sxs-lookup"><span data-stu-id="0f62e-1051">function</span></span>||<span data-ttu-id="0f62e-1052">メソッドが完了すると、`callback` パラメーターに渡された関数が、[`asyncResult`](/javascript/api/office/office.asyncresult) オブジェクトである 1 つのパラメーター `AsyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="0f62e-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f62e-1053">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-1053">Requirements</span></span>

|<span data-ttu-id="0f62e-1054">要件</span><span class="sxs-lookup"><span data-stu-id="0f62e-1054">Requirement</span></span>| <span data-ttu-id="0f62e-1055">値</span><span class="sxs-lookup"><span data-stu-id="0f62e-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f62e-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0f62e-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f62e-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="0f62e-1057">1.2</span></span>|
|[<span data-ttu-id="0f62e-1058">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0f62e-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0f62e-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0f62e-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="0f62e-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0f62e-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0f62e-1061">新規作成</span><span class="sxs-lookup"><span data-stu-id="0f62e-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0f62e-1062">例</span><span class="sxs-lookup"><span data-stu-id="0f62e-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```