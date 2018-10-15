
# <a name="item"></a><span data-ttu-id="f0777-101">項目</span><span class="sxs-lookup"><span data-stu-id="f0777-101">item</span></span>

### <span data-ttu-id="f0777-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f0777-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="f0777-p102">`item` 名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予約にアクセスします。`item` の種類を [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) プロパティを使用して決定できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-106">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-106">Requirements</span></span>

|<span data-ttu-id="f0777-107">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-107">Requirement</span></span>| <span data-ttu-id="f0777-108">値</span><span class="sxs-lookup"><span data-stu-id="f0777-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-110">1.0</span></span>|
|[<span data-ttu-id="f0777-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="f0777-112">Restricted</span></span>|
|[<span data-ttu-id="f0777-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-114">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="f0777-115">例</span><span class="sxs-lookup"><span data-stu-id="f0777-115">Example</span></span>

<span data-ttu-id="f0777-116">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f0777-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f0777-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="f0777-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="f0777-118">添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f0777-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="f0777-p103">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-121">潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="f0777-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f0777-122">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="f0777-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-123">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-123">Type:</span></span>

*   <span data-ttu-id="f0777-124">配列.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f0777-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-125">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-125">Requirements</span></span>

|<span data-ttu-id="f0777-126">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-126">Requirement</span></span>| <span data-ttu-id="f0777-127">値</span><span class="sxs-lookup"><span data-stu-id="f0777-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-128">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-129">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-129">1.0</span></span>|
|[<span data-ttu-id="f0777-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-131">ReadItem</span></span>|
|[<span data-ttu-id="f0777-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-134">例</span><span class="sxs-lookup"><span data-stu-id="f0777-134">Example</span></span>

<span data-ttu-id="f0777-135">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="f0777-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="f0777-136">bcc:[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-136">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="f0777-137">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f0777-138">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f0777-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-139">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-139">Type:</span></span>

*   [<span data-ttu-id="f0777-140">受信者</span><span class="sxs-lookup"><span data-stu-id="f0777-140">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f0777-141">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-141">Requirements</span></span>

|<span data-ttu-id="f0777-142">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-142">Requirement</span></span>| <span data-ttu-id="f0777-143">値</span><span class="sxs-lookup"><span data-stu-id="f0777-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-145">1.1</span><span class="sxs-lookup"><span data-stu-id="f0777-145">1.1</span></span>|
|[<span data-ttu-id="f0777-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-147">ReadItem</span></span>|
|[<span data-ttu-id="f0777-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-149">新規作成</span><span class="sxs-lookup"><span data-stu-id="f0777-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-150">例</span><span class="sxs-lookup"><span data-stu-id="f0777-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="f0777-151">本文:[本文](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="f0777-151">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="f0777-152">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-153">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-153">Type:</span></span>

*   [<span data-ttu-id="f0777-154">本文</span><span class="sxs-lookup"><span data-stu-id="f0777-154">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="f0777-155">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-155">Requirements</span></span>

|<span data-ttu-id="f0777-156">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-156">Requirement</span></span>| <span data-ttu-id="f0777-157">値</span><span class="sxs-lookup"><span data-stu-id="f0777-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-158">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f0777-159">1.1</span></span>|
|[<span data-ttu-id="f0777-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-161">ReadItem</span></span>|
|[<span data-ttu-id="f0777-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-163">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="f0777-164">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="f0777-165">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0777-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f0777-166">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0777-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-167">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-167">Read mode</span></span>

<span data-ttu-id="f0777-p107">`cc` プロパティは、メッセージの **CC** 行にある各受信者について、`EmailAddressDetails` オブ ジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-170">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-170">Compose mode</span></span>

<span data-ttu-id="f0777-171">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-172">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-172">Type:</span></span>

*   <span data-ttu-id="f0777-173">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-174">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-174">Requirements</span></span>

|<span data-ttu-id="f0777-175">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-175">Requirement</span></span>| <span data-ttu-id="f0777-176">値</span><span class="sxs-lookup"><span data-stu-id="f0777-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-177">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-178">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-178">1.0</span></span>|
|[<span data-ttu-id="f0777-179">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-180">ReadItem</span></span>|
|[<span data-ttu-id="f0777-181">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-182">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-183">例</span><span class="sxs-lookup"><span data-stu-id="f0777-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f0777-184">（空白が可能）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="f0777-185">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f0777-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="f0777-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f0777-p109">作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-190">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-190">Type:</span></span>

*   <span data-ttu-id="f0777-191">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-192">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-192">Requirements</span></span>

|<span data-ttu-id="f0777-193">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-193">Requirement</span></span>| <span data-ttu-id="f0777-194">値</span><span class="sxs-lookup"><span data-stu-id="f0777-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-195">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-196">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-196">1.0</span></span>|
|[<span data-ttu-id="f0777-197">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-198">ReadItem</span></span>|
|[<span data-ttu-id="f0777-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-200">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f0777-201">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="f0777-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="f0777-p110">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-204">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-204">Type:</span></span>

*   <span data-ttu-id="f0777-205">日付</span><span class="sxs-lookup"><span data-stu-id="f0777-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-206">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-206">Requirements</span></span>

|<span data-ttu-id="f0777-207">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-207">Requirement</span></span>| <span data-ttu-id="f0777-208">値</span><span class="sxs-lookup"><span data-stu-id="f0777-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-210">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-210">1.0</span></span>|
|[<span data-ttu-id="f0777-211">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-212">ReadItem</span></span>|
|[<span data-ttu-id="f0777-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-214">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-215">例</span><span class="sxs-lookup"><span data-stu-id="f0777-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f0777-216">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="f0777-216">dateTimeModified :Date</span></span>

<span data-ttu-id="f0777-p111">アイテムが最後に変更された日時を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-219">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-220">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-220">Type:</span></span>

*   <span data-ttu-id="f0777-221">日付</span><span class="sxs-lookup"><span data-stu-id="f0777-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-222">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-222">Requirements</span></span>

|<span data-ttu-id="f0777-223">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-223">Requirement</span></span>| <span data-ttu-id="f0777-224">値</span><span class="sxs-lookup"><span data-stu-id="f0777-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-226">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-226">1.0</span></span>|
|[<span data-ttu-id="f0777-227">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-228">ReadItem</span></span>|
|[<span data-ttu-id="f0777-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-230">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-231">例</span><span class="sxs-lookup"><span data-stu-id="f0777-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="f0777-232">end:日付|[時間](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="f0777-232">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="f0777-233">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f0777-p112">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f0777-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-236">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-236">Read mode</span></span>

<span data-ttu-id="f0777-237">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-238">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-238">Compose mode</span></span>

<span data-ttu-id="f0777-239">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f0777-240">[`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0777-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-241">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-241">Type:</span></span>

*   <span data-ttu-id="f0777-242">日付| [時間](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="f0777-242">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-243">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-243">Requirements</span></span>

|<span data-ttu-id="f0777-244">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-244">Requirement</span></span>| <span data-ttu-id="f0777-245">値</span><span class="sxs-lookup"><span data-stu-id="f0777-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-246">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-247">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-247">1.0</span></span>|
|[<span data-ttu-id="f0777-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-249">ReadItem</span></span>|
|[<span data-ttu-id="f0777-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-251">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-252">例</span><span class="sxs-lookup"><span data-stu-id="f0777-252">Example</span></span>

<span data-ttu-id="f0777-253">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="f0777-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f0777-254">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="f0777-p113">メッセージの送信者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f0777-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-259">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="f0777-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-260">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-260">Type:</span></span>

*   [<span data-ttu-id="f0777-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0777-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f0777-262">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-262">Requirements</span></span>

|<span data-ttu-id="f0777-263">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-263">Requirement</span></span>| <span data-ttu-id="f0777-264">値</span><span class="sxs-lookup"><span data-stu-id="f0777-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-265">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-266">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-266">1.0</span></span>|
|[<span data-ttu-id="f0777-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-268">ReadItem</span></span>|
|[<span data-ttu-id="f0777-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-270">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f0777-271">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-271">internetMessageId :String</span></span>

<span data-ttu-id="f0777-p115">電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-274">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-274">Type:</span></span>

*   <span data-ttu-id="f0777-275">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-276">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-276">Requirements</span></span>

|<span data-ttu-id="f0777-277">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-277">Requirement</span></span>| <span data-ttu-id="f0777-278">値</span><span class="sxs-lookup"><span data-stu-id="f0777-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-279">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-280">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-280">1.0</span></span>|
|[<span data-ttu-id="f0777-281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-282">ReadItem</span></span>|
|[<span data-ttu-id="f0777-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-284">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-285">例</span><span class="sxs-lookup"><span data-stu-id="f0777-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f0777-286">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-286">itemClass :String</span></span>

<span data-ttu-id="f0777-p116">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f0777-p117">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f0777-291">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-291">Type</span></span> | <span data-ttu-id="f0777-292">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-292">Description</span></span> | <span data-ttu-id="f0777-293">項目のクラス</span><span class="sxs-lookup"><span data-stu-id="f0777-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f0777-294">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="f0777-294">Appointment items</span></span> | <span data-ttu-id="f0777-295">これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="f0777-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="f0777-296">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="f0777-296">Message items</span></span> | <span data-ttu-id="f0777-297">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0777-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f0777-298">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-299">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-299">Type:</span></span>

*   <span data-ttu-id="f0777-300">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-301">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-301">Requirements</span></span>

|<span data-ttu-id="f0777-302">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-302">Requirement</span></span>| <span data-ttu-id="f0777-303">値</span><span class="sxs-lookup"><span data-stu-id="f0777-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-304">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-305">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-305">1.0</span></span>|
|[<span data-ttu-id="f0777-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-307">ReadItem</span></span>|
|[<span data-ttu-id="f0777-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-310">例</span><span class="sxs-lookup"><span data-stu-id="f0777-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f0777-311">（空白が可能） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-311">(nullable) itemId :String</span></span>

<span data-ttu-id="f0777-p118">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-314">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="f0777-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f0777-315">`itemId` プロパティは、Outlook Entry ID または Outlook REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="f0777-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f0777-316">この値を使用して REST API の呼び出しを行う前に、 必要な設定1.3より利用可能な `Office.context.mailbox.convertToRestId`を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0777-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="f0777-317">詳細については、[「Outlook アドインから Outlook REST API の使用」](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="f0777-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-318">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-318">Type:</span></span>

*   <span data-ttu-id="f0777-319">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-320">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-320">Requirements</span></span>

|<span data-ttu-id="f0777-321">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-321">Requirement</span></span>| <span data-ttu-id="f0777-322">値</span><span class="sxs-lookup"><span data-stu-id="f0777-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-323">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-324">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-324">1.0</span></span>|
|[<span data-ttu-id="f0777-325">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-326">ReadItem</span></span>|
|[<span data-ttu-id="f0777-327">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-328">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-329">例</span><span class="sxs-lookup"><span data-stu-id="f0777-329">Example</span></span>

<span data-ttu-id="f0777-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="f0777-332">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f0777-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f0777-333">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f0777-334">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="f0777-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-335">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-335">Type:</span></span>

*   [<span data-ttu-id="f0777-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f0777-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f0777-337">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-337">Requirements</span></span>

|<span data-ttu-id="f0777-338">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-338">Requirement</span></span>| <span data-ttu-id="f0777-339">値</span><span class="sxs-lookup"><span data-stu-id="f0777-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-340">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-341">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-341">1.0</span></span>|
|[<span data-ttu-id="f0777-342">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-343">ReadItem</span></span>|
|[<span data-ttu-id="f0777-344">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-345">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-346">例</span><span class="sxs-lookup"><span data-stu-id="f0777-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="f0777-347">位置: 文字列|[](/javascript/api/outlook_1_1/office.location)位置</span><span class="sxs-lookup"><span data-stu-id="f0777-347">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="f0777-348">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-349">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-349">Read mode</span></span>

<span data-ttu-id="f0777-350">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-351">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-351">Compose mode</span></span>

<span data-ttu-id="f0777-352">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-353">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-353">Type:</span></span>

*   <span data-ttu-id="f0777-354">文字列 | [場所](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="f0777-354">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-355">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-355">Requirements</span></span>

|<span data-ttu-id="f0777-356">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-356">Requirement</span></span>| <span data-ttu-id="f0777-357">値</span><span class="sxs-lookup"><span data-stu-id="f0777-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-358">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-359">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-359">1.0</span></span>|
|[<span data-ttu-id="f0777-360">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-361">ReadItem</span></span>|
|[<span data-ttu-id="f0777-362">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-363">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-364">例</span><span class="sxs-lookup"><span data-stu-id="f0777-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f0777-365">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-365">normalizedSubject :String</span></span>

<span data-ttu-id="f0777-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f0777-p122">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-370">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-370">Type:</span></span>

*   <span data-ttu-id="f0777-371">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-372">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-372">Requirements</span></span>

|<span data-ttu-id="f0777-373">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-373">Requirement</span></span>| <span data-ttu-id="f0777-374">値</span><span class="sxs-lookup"><span data-stu-id="f0777-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-375">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-376">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-376">1.0</span></span>|
|[<span data-ttu-id="f0777-377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-378">ReadItem</span></span>|
|[<span data-ttu-id="f0777-379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-380">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-381">例</span><span class="sxs-lookup"><span data-stu-id="f0777-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="f0777-382">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="f0777-383">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0777-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f0777-384">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0777-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-385">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-385">Read mode</span></span>

<span data-ttu-id="f0777-386">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-387">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-387">Compose mode</span></span>

<span data-ttu-id="f0777-388">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-389">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-389">Type:</span></span>

*   <span data-ttu-id="f0777-390">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-391">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-391">Requirements</span></span>

|<span data-ttu-id="f0777-392">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-392">Requirement</span></span>| <span data-ttu-id="f0777-393">値</span><span class="sxs-lookup"><span data-stu-id="f0777-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-394">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-395">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-395">1.0</span></span>|
|[<span data-ttu-id="f0777-396">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-397">ReadItem</span></span>|
|[<span data-ttu-id="f0777-398">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-399">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-400">例</span><span class="sxs-lookup"><span data-stu-id="f0777-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="f0777-401">開催者:[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f0777-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="f0777-p124">指定の会議の開催者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-404">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-404">Type:</span></span>

*   [<span data-ttu-id="f0777-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0777-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f0777-406">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-406">Requirements</span></span>

|<span data-ttu-id="f0777-407">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-407">Requirement</span></span>| <span data-ttu-id="f0777-408">値</span><span class="sxs-lookup"><span data-stu-id="f0777-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-409">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-410">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-410">1.0</span></span>|
|[<span data-ttu-id="f0777-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-412">ReadItem</span></span>|
|[<span data-ttu-id="f0777-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-415">例</span><span class="sxs-lookup"><span data-stu-id="f0777-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="f0777-416">requiredAttendees: 配列 。<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="f0777-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="f0777-417">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0777-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f0777-418">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0777-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-419">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-419">Read mode</span></span>

<span data-ttu-id="f0777-420">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-421">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-421">Compose mode</span></span>

<span data-ttu-id="f0777-422">`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-423">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-423">Type:</span></span>

*   <span data-ttu-id="f0777-424">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-425">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-425">Requirements</span></span>

|<span data-ttu-id="f0777-426">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-426">Requirement</span></span>| <span data-ttu-id="f0777-427">値</span><span class="sxs-lookup"><span data-stu-id="f0777-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-428">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-429">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-429">1.0</span></span>|
|[<span data-ttu-id="f0777-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-431">ReadItem</span></span>|
|[<span data-ttu-id="f0777-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-433">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-434">例</span><span class="sxs-lookup"><span data-stu-id="f0777-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="f0777-435">送信者:[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f0777-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="f0777-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f0777-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-440">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="f0777-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-441">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-441">Type:</span></span>

*   [<span data-ttu-id="f0777-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f0777-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f0777-443">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-443">Requirements</span></span>

|<span data-ttu-id="f0777-444">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-444">Requirement</span></span>| <span data-ttu-id="f0777-445">値</span><span class="sxs-lookup"><span data-stu-id="f0777-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-446">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-447">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-447">1.0</span></span>|
|[<span data-ttu-id="f0777-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-449">ReadItem</span></span>|
|[<span data-ttu-id="f0777-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-451">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-452">例</span><span class="sxs-lookup"><span data-stu-id="f0777-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="f0777-453">開始:日付| [時間](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="f0777-453">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="f0777-454">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f0777-p128">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f0777-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-457">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-457">Read mode</span></span>

<span data-ttu-id="f0777-458">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-459">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-459">Compose mode</span></span>

<span data-ttu-id="f0777-460">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f0777-461">[ `Time.setAsync` ](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f0777-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-462">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-462">Type:</span></span>

*   <span data-ttu-id="f0777-463">日付| [時間](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="f0777-463">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-464">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-464">Requirements</span></span>

|<span data-ttu-id="f0777-465">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-465">Requirement</span></span>| <span data-ttu-id="f0777-466">値</span><span class="sxs-lookup"><span data-stu-id="f0777-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-467">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-468">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-468">1.0</span></span>|
|[<span data-ttu-id="f0777-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-470">ReadItem</span></span>|
|[<span data-ttu-id="f0777-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-472">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-473">例</span><span class="sxs-lookup"><span data-stu-id="f0777-473">Example</span></span>

<span data-ttu-id="f0777-474">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="f0777-475">件名: 文字列 | [件名](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f0777-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="f0777-476">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f0777-477">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f0777-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-478">Read mode</span></span>

<span data-ttu-id="f0777-p129">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f0777-481">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-481">Compose mode</span></span>

<span data-ttu-id="f0777-482">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f0777-483">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-483">Type:</span></span>

*   <span data-ttu-id="f0777-484">文字列 | [件名](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f0777-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-485">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-485">Requirements</span></span>

|<span data-ttu-id="f0777-486">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-486">Requirement</span></span>| <span data-ttu-id="f0777-487">値</span><span class="sxs-lookup"><span data-stu-id="f0777-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-488">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-489">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-489">1.0</span></span>|
|[<span data-ttu-id="f0777-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-491">ReadItem</span></span>|
|[<span data-ttu-id="f0777-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-493">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="f0777-494">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="f0777-495">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f0777-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f0777-496">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0777-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f0777-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f0777-497">Read mode</span></span>

<span data-ttu-id="f0777-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f0777-500">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f0777-500">Compose mode</span></span>

<span data-ttu-id="f0777-501">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f0777-502">種類:</span><span class="sxs-lookup"><span data-stu-id="f0777-502">Type:</span></span>

*   <span data-ttu-id="f0777-503">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f0777-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-504">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-504">Requirements</span></span>

|<span data-ttu-id="f0777-505">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-505">Requirement</span></span>| <span data-ttu-id="f0777-506">値</span><span class="sxs-lookup"><span data-stu-id="f0777-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-507">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-508">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-508">1.0</span></span>|
|[<span data-ttu-id="f0777-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-510">ReadItem</span></span>|
|[<span data-ttu-id="f0777-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-512">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-513">例</span><span class="sxs-lookup"><span data-stu-id="f0777-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f0777-514">メソッド</span><span class="sxs-lookup"><span data-stu-id="f0777-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f0777-515">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="f0777-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f0777-516">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f0777-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f0777-517">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="f0777-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f0777-518">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-519">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-519">Parameters:</span></span>

|<span data-ttu-id="f0777-520">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-520">Name</span></span>| <span data-ttu-id="f0777-521">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-521">Type</span></span>| <span data-ttu-id="f0777-522">属性</span><span class="sxs-lookup"><span data-stu-id="f0777-522">Attributes</span></span>| <span data-ttu-id="f0777-523">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f0777-524">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-524">String</span></span>||<span data-ttu-id="f0777-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="f0777-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f0777-527">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-527">String</span></span>||<span data-ttu-id="f0777-p133">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="f0777-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f0777-530">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-530">Object</span></span>| <span data-ttu-id="f0777-531">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-531">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-532">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f0777-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0777-533">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-533">Object</span></span>| <span data-ttu-id="f0777-534">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-534">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-535">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0777-536">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-536">function</span></span>| <span data-ttu-id="f0777-537">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-537">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-538">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f0777-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0777-539">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f0777-540">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0777-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0777-541">エラー</span><span class="sxs-lookup"><span data-stu-id="f0777-541">Errors</span></span>

| <span data-ttu-id="f0777-542">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0777-542">Error code</span></span> | <span data-ttu-id="f0777-543">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f0777-544">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="f0777-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f0777-545">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="f0777-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f0777-546">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f0777-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0777-547">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-547">Requirements</span></span>

|<span data-ttu-id="f0777-548">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-548">Requirement</span></span>| <span data-ttu-id="f0777-549">値</span><span class="sxs-lookup"><span data-stu-id="f0777-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-550">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-551">1.1</span><span class="sxs-lookup"><span data-stu-id="f0777-551">1.1</span></span>|
|[<span data-ttu-id="f0777-552">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0777-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0777-554">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-555">新規作成</span><span class="sxs-lookup"><span data-stu-id="f0777-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-556">例</span><span class="sxs-lookup"><span data-stu-id="f0777-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f0777-557">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="f0777-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f0777-558">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f0777-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f0777-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="f0777-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f0777-562">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f0777-563">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-564">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-564">Parameters:</span></span>

|<span data-ttu-id="f0777-565">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-565">Name</span></span>| <span data-ttu-id="f0777-566">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-566">Type</span></span>| <span data-ttu-id="f0777-567">属性</span><span class="sxs-lookup"><span data-stu-id="f0777-567">Attributes</span></span>| <span data-ttu-id="f0777-568">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f0777-569">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-569">String</span></span>||<span data-ttu-id="f0777-p135">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f0777-572">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-572">String</span></span>||<span data-ttu-id="f0777-p136">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f0777-575">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-575">Object</span></span>| <span data-ttu-id="f0777-576">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-576">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-577">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f0777-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0777-578">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-578">Object</span></span>| <span data-ttu-id="f0777-579">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-579">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-580">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0777-581">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-581">function</span></span>| <span data-ttu-id="f0777-582">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-582">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-583">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f0777-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0777-584">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f0777-585">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0777-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0777-586">エラー</span><span class="sxs-lookup"><span data-stu-id="f0777-586">Errors</span></span>

| <span data-ttu-id="f0777-587">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0777-587">Error code</span></span> | <span data-ttu-id="f0777-588">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f0777-589">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f0777-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0777-590">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-590">Requirements</span></span>

|<span data-ttu-id="f0777-591">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-591">Requirement</span></span>| <span data-ttu-id="f0777-592">値</span><span class="sxs-lookup"><span data-stu-id="f0777-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-593">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-594">1.1</span><span class="sxs-lookup"><span data-stu-id="f0777-594">1.1</span></span>|
|[<span data-ttu-id="f0777-595">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0777-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0777-597">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-598">新規作成</span><span class="sxs-lookup"><span data-stu-id="f0777-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-599">例</span><span class="sxs-lookup"><span data-stu-id="f0777-599">Example</span></span>

<span data-ttu-id="f0777-600">次の例では、既存の Outlook 項目を名前付き `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f0777-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f0777-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f0777-602">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-603">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f0777-604">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f0777-605">文字列パラメータのいずれかが制限値を超えると、`displayReplyAllForm` は例外を提案します。</span><span class="sxs-lookup"><span data-stu-id="f0777-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-606">`displayReplyAllForm` に対する呼び出しに添付ファイルを含める機能は、設定要件 である version 1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-606">NOTE: The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="f0777-607">添付ファイルのサポートは、設定要件が version 1.2 以降で `displayReplyAllForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0777-607">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-608">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-608">Parameters:</span></span>

|<span data-ttu-id="f0777-609">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-609">Name</span></span>| <span data-ttu-id="f0777-610">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-610">Type</span></span>| <span data-ttu-id="f0777-611">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f0777-612">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-612">String &#124; Object</span></span>| |<span data-ttu-id="f0777-p138">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f0777-615">**または**</span><span class="sxs-lookup"><span data-stu-id="f0777-615">**OR**</span></span><br/><span data-ttu-id="f0777-p139">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f0777-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f0777-618">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-618">String</span></span> | <span data-ttu-id="f0777-619">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-619">&lt;optional&gt;</span></span> | <span data-ttu-id="f0777-p140">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="f0777-622">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-622">function</span></span> | <span data-ttu-id="f0777-623">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-623">&lt;optional&gt;</span></span> | <span data-ttu-id="f0777-624">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-624">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0777-625">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-625">Requirements</span></span>

|<span data-ttu-id="f0777-626">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-626">Requirement</span></span>| <span data-ttu-id="f0777-627">値</span><span class="sxs-lookup"><span data-stu-id="f0777-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-628">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-629">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-629">1.0</span></span>|
|[<span data-ttu-id="f0777-630">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-631">ReadItem</span></span>|
|[<span data-ttu-id="f0777-632">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-633">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-633">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0777-634">例</span><span class="sxs-lookup"><span data-stu-id="f0777-634">Examples</span></span>

<span data-ttu-id="f0777-635">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f0777-635">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f0777-636">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-636">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f0777-637">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-637">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f0777-638">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-638">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f0777-639">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f0777-639">displayReplyForm(formData)</span></span>

<span data-ttu-id="f0777-640">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-640">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-641">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-641">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f0777-642">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-642">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f0777-643">文字列パラメータのいずれかが制限値を超えると、`displayReplyForm` は例外を提案します。</span><span class="sxs-lookup"><span data-stu-id="f0777-643">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-644">`displayReplyForm` に対する呼び出しに添付ファイルを含める機能は、設定要件 である version 1.1 ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-644">NOTE: The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="f0777-645">添付ファイルのサポートは、設定要件が version 1.2 以降で `displayReplyForm` に追加されました。</span><span class="sxs-lookup"><span data-stu-id="f0777-645">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-646">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-646">Parameters:</span></span>

|<span data-ttu-id="f0777-647">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-647">Name</span></span>| <span data-ttu-id="f0777-648">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-648">Type</span></span>| <span data-ttu-id="f0777-649">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-649">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f0777-650">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-650">String &#124; Object</span></span>| | <span data-ttu-id="f0777-p142">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f0777-653">**または**</span><span class="sxs-lookup"><span data-stu-id="f0777-653">**OR**</span></span><br/><span data-ttu-id="f0777-p143">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f0777-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f0777-656">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-656">String</span></span> | <span data-ttu-id="f0777-657">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-657">&lt;optional&gt;</span></span> | <span data-ttu-id="f0777-p144">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f0777-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="f0777-660">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-660">function</span></span> | <span data-ttu-id="f0777-661">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-661">&lt;optional&gt;</span></span> | <span data-ttu-id="f0777-662">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-662">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0777-663">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-663">Requirements</span></span>

|<span data-ttu-id="f0777-664">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-664">Requirement</span></span>| <span data-ttu-id="f0777-665">値</span><span class="sxs-lookup"><span data-stu-id="f0777-665">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-666">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-666">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-667">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-667">1.0</span></span>|
|[<span data-ttu-id="f0777-668">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-668">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-669">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-669">ReadItem</span></span>|
|[<span data-ttu-id="f0777-670">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-670">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-671">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-671">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f0777-672">例</span><span class="sxs-lookup"><span data-stu-id="f0777-672">Examples</span></span>

<span data-ttu-id="f0777-673">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f0777-673">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f0777-674">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-674">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f0777-675">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-675">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f0777-676">本文とコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f0777-676">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="f0777-677">getEntities() → {[エンティティ](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f0777-677">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="f0777-678">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-678">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-679">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-679">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-680">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-680">Requirements</span></span>

|<span data-ttu-id="f0777-681">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-681">Requirement</span></span>| <span data-ttu-id="f0777-682">値</span><span class="sxs-lookup"><span data-stu-id="f0777-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-683">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-684">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-684">1.0</span></span>|
|[<span data-ttu-id="f0777-685">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-685">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-686">ReadItem</span></span>|
|[<span data-ttu-id="f0777-687">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-687">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-688">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-688">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0777-689">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f0777-689">Returns:</span></span>

<span data-ttu-id="f0777-690">種類: [エンティティ](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f0777-690">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f0777-691">例</span><span class="sxs-lookup"><span data-stu-id="f0777-691">Example</span></span>

<span data-ttu-id="f0777-692">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f0777-692">The following example accesses the contacts entities on the current item.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="f0777-693">getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="f0777-693">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f0777-694">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f0777-694">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-695">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-695">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-696">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-696">Parameters:</span></span>

|<span data-ttu-id="f0777-697">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-697">Name</span></span>| <span data-ttu-id="f0777-698">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-698">Type</span></span>| <span data-ttu-id="f0777-699">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-699">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f0777-700">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f0777-700">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="f0777-701">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="f0777-701">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0777-702">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-702">Requirements</span></span>

|<span data-ttu-id="f0777-703">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-703">Requirement</span></span>| <span data-ttu-id="f0777-704">値</span><span class="sxs-lookup"><span data-stu-id="f0777-704">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-705">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-705">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-706">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-706">1.0</span></span>|
|[<span data-ttu-id="f0777-707">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-707">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-708">制限あり</span><span class="sxs-lookup"><span data-stu-id="f0777-708">Restricted</span></span>|
|[<span data-ttu-id="f0777-709">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-709">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-710">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-710">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0777-711">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f0777-711">Returns:</span></span>

<span data-ttu-id="f0777-712">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-712">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f0777-713">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-713">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="f0777-714">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f0777-714">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f0777-715">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="f0777-715">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f0777-716">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="f0777-716">Value of `entityType`</span></span> | <span data-ttu-id="f0777-717">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="f0777-717">Type of objects in returned array</span></span> | <span data-ttu-id="f0777-718">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="f0777-718">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f0777-719">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-719">String</span></span> | <span data-ttu-id="f0777-720">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0777-720">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f0777-721">連絡先</span><span class="sxs-lookup"><span data-stu-id="f0777-721">Contact</span></span> | <span data-ttu-id="f0777-722">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0777-722">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f0777-723">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-723">String</span></span> | <span data-ttu-id="f0777-724">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0777-724">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f0777-725">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f0777-725">MeetingSuggestion</span></span> | <span data-ttu-id="f0777-726">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0777-726">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f0777-727">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f0777-727">PhoneNumber</span></span> | <span data-ttu-id="f0777-728">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0777-728">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f0777-729">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f0777-729">TaskSuggestion</span></span> | <span data-ttu-id="f0777-730">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f0777-730">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f0777-731">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-731">String</span></span> | <span data-ttu-id="f0777-732">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f0777-732">**Restricted**</span></span> |

<span data-ttu-id="f0777-733">種類:配列.<(文字列|[連絡先](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f0777-733">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="f0777-734">例</span><span class="sxs-lookup"><span data-stu-id="f0777-734">Example</span></span>

<span data-ttu-id="f0777-735">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f0777-735">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="f0777-736">getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f0777-736">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f0777-737">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-737">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-738">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-738">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f0777-739">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-739">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-740">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-740">Parameters:</span></span>

|<span data-ttu-id="f0777-741">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-741">Name</span></span>| <span data-ttu-id="f0777-742">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-742">Type</span></span>| <span data-ttu-id="f0777-743">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-743">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f0777-744">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-744">String</span></span>|<span data-ttu-id="f0777-745">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="f0777-745">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0777-746">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-746">Requirements</span></span>

|<span data-ttu-id="f0777-747">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-747">Requirement</span></span>| <span data-ttu-id="f0777-748">値</span><span class="sxs-lookup"><span data-stu-id="f0777-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-749">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-750">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-750">1.0</span></span>|
|[<span data-ttu-id="f0777-751">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-752">ReadItem</span></span>|
|[<span data-ttu-id="f0777-753">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-754">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0777-755">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f0777-755">Returns:</span></span>

<span data-ttu-id="f0777-p146">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="f0777-758">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f0777-758">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="f0777-759">getRegExMatches() → {オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f0777-759">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f0777-760">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-760">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-761">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-761">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f0777-p147">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f0777-765">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f0777-765">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f0777-766">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="f0777-766">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="f0777-p148">項目の本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f0777-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0777-769">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-769">Requirements</span></span>

|<span data-ttu-id="f0777-770">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-770">Requirement</span></span>| <span data-ttu-id="f0777-771">値</span><span class="sxs-lookup"><span data-stu-id="f0777-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-772">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-773">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-773">1.0</span></span>|
|[<span data-ttu-id="f0777-774">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-774">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-775">ReadItem</span></span>|
|[<span data-ttu-id="f0777-776">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-776">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-777">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-777">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0777-778">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f0777-778">Returns:</span></span>

<span data-ttu-id="f0777-p149">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f0777-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f0777-781">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="f0777-781">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f0777-782">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-782">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f0777-783">例</span><span class="sxs-lookup"><span data-stu-id="f0777-783">Example</span></span>

<span data-ttu-id="f0777-784">次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="f0777-784">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f0777-785">getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="f0777-785">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f0777-786">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f0777-786">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f0777-787">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f0777-787">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f0777-788">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="f0777-788">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f0777-p150">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f0777-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-791">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-791">Parameters:</span></span>

|<span data-ttu-id="f0777-792">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-792">Name</span></span>| <span data-ttu-id="f0777-793">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-793">Type</span></span>| <span data-ttu-id="f0777-794">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-794">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f0777-795">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-795">String</span></span>|<span data-ttu-id="f0777-796">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="f0777-796">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0777-797">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-797">Requirements</span></span>

|<span data-ttu-id="f0777-798">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-798">Requirement</span></span>| <span data-ttu-id="f0777-799">値</span><span class="sxs-lookup"><span data-stu-id="f0777-799">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-800">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-800">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-801">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-801">1.0</span></span>|
|[<span data-ttu-id="f0777-802">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-802">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-803">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-803">ReadItem</span></span>|
|[<span data-ttu-id="f0777-804">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-804">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-805">読み取り</span><span class="sxs-lookup"><span data-stu-id="f0777-805">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f0777-806">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f0777-806">Returns:</span></span>

<span data-ttu-id="f0777-807">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="f0777-807">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f0777-808">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="f0777-808">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f0777-809">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="f0777-809">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f0777-810">例</span><span class="sxs-lookup"><span data-stu-id="f0777-810">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f0777-811">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f0777-811">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f0777-812">選択された項目のアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f0777-812">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f0777-p151">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="f0777-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-816">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-816">Parameters:</span></span>

|<span data-ttu-id="f0777-817">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-817">Name</span></span>| <span data-ttu-id="f0777-818">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-818">Type</span></span>| <span data-ttu-id="f0777-819">属性</span><span class="sxs-lookup"><span data-stu-id="f0777-819">Attributes</span></span>| <span data-ttu-id="f0777-820">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-820">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f0777-821">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-821">function</span></span>||<span data-ttu-id="f0777-822">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f0777-822">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f0777-823">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="f0777-823">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f0777-824">このオブジェクトは、項目からカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="f0777-824">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f0777-825">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-825">Object</span></span>| <span data-ttu-id="f0777-826">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-826">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-827">開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-827">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="f0777-828">このオブジェクトは、コールバック関数の`asyncResult.asyncContext`プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f0777-828">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0777-829">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-829">Requirements</span></span>

|<span data-ttu-id="f0777-830">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-830">Requirement</span></span>| <span data-ttu-id="f0777-831">値</span><span class="sxs-lookup"><span data-stu-id="f0777-831">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-832">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-832">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-833">1.0</span><span class="sxs-lookup"><span data-stu-id="f0777-833">1.0</span></span>|
|[<span data-ttu-id="f0777-834">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-834">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-835">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f0777-835">ReadItem</span></span>|
|[<span data-ttu-id="f0777-836">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-836">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-837">作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f0777-837">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-838">例</span><span class="sxs-lookup"><span data-stu-id="f0777-838">Example</span></span>

<span data-ttu-id="f0777-p154">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在の項目に固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、コード サンプルは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f0777-842">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="f0777-842">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f0777-843">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f0777-843">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f0777-p155">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="f0777-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f0777-848">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f0777-848">Parameters:</span></span>

|<span data-ttu-id="f0777-849">名前</span><span class="sxs-lookup"><span data-stu-id="f0777-849">Name</span></span>| <span data-ttu-id="f0777-850">種類</span><span class="sxs-lookup"><span data-stu-id="f0777-850">Type</span></span>| <span data-ttu-id="f0777-851">属性</span><span class="sxs-lookup"><span data-stu-id="f0777-851">Attributes</span></span>| <span data-ttu-id="f0777-852">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-852">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f0777-853">文字列</span><span class="sxs-lookup"><span data-stu-id="f0777-853">String</span></span>||<span data-ttu-id="f0777-p156">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f0777-p156">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="f0777-856">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-856">Object</span></span>| <span data-ttu-id="f0777-857">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-857">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-858">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f0777-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f0777-859">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f0777-859">Object</span></span>| <span data-ttu-id="f0777-860">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-860">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-861">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f0777-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f0777-862">関数</span><span class="sxs-lookup"><span data-stu-id="f0777-862">function</span></span>| <span data-ttu-id="f0777-863">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f0777-863">&lt;optional&gt;</span></span>|<span data-ttu-id="f0777-864">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f0777-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f0777-865">添付ファイルの削除に失敗すると、`asyncResult.error`プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f0777-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f0777-866">エラー</span><span class="sxs-lookup"><span data-stu-id="f0777-866">Errors</span></span>

| <span data-ttu-id="f0777-867">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f0777-867">Error code</span></span> | <span data-ttu-id="f0777-868">説明</span><span class="sxs-lookup"><span data-stu-id="f0777-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f0777-869">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="f0777-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0777-870">要件</span><span class="sxs-lookup"><span data-stu-id="f0777-870">Requirements</span></span>

|<span data-ttu-id="f0777-871">必要条件</span><span class="sxs-lookup"><span data-stu-id="f0777-871">Requirement</span></span>| <span data-ttu-id="f0777-872">値</span><span class="sxs-lookup"><span data-stu-id="f0777-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0777-873">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f0777-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0777-874">1.1</span><span class="sxs-lookup"><span data-stu-id="f0777-874">1.1</span></span>|
|[<span data-ttu-id="f0777-875">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f0777-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f0777-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f0777-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="f0777-877">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f0777-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f0777-878">新規作成</span><span class="sxs-lookup"><span data-stu-id="f0777-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f0777-879">例</span><span class="sxs-lookup"><span data-stu-id="f0777-879">Example</span></span>

<span data-ttu-id="f0777-880">次のコードは、「0」の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f0777-880">The following code removes an attachment with an identifier of '0'.</span></span>

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