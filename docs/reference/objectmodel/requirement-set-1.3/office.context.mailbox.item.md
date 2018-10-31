
# <a name="item"></a><span data-ttu-id="f8455-101">項目</span><span class="sxs-lookup"><span data-stu-id="f8455-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f8455-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f8455-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f8455-p101">`item` の名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予定にアクセスします。[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-105">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-105">Requirements</span></span>

|<span data-ttu-id="f8455-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-106">Requirement</span></span>| <span data-ttu-id="f8455-107">値</span><span class="sxs-lookup"><span data-stu-id="f8455-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-108">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-109">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-109">1.0</span></span>|
|[<span data-ttu-id="f8455-110">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="f8455-111">Restricted</span></span>|
|[<span data-ttu-id="f8455-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-113">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="f8455-114">例</span><span class="sxs-lookup"><span data-stu-id="f8455-114">Example</span></span>

<span data-ttu-id="f8455-115">次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="f8455-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="f8455-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="f8455-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="f8455-117">添付ファイル：配列.<[添付ファイルの詳細 ](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f8455-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="f8455-p102">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-120">潜在的なセキュリティ問題により特定の種類のファイルは、Outlookでブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="f8455-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f8455-121">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="f8455-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-122">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-122">Type:</span></span>

*   <span data-ttu-id="f8455-123">配列.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f8455-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-124">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-124">Requirements</span></span>

|<span data-ttu-id="f8455-125">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-125">Requirement</span></span>| <span data-ttu-id="f8455-126">値</span><span class="sxs-lookup"><span data-stu-id="f8455-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-127">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-128">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-128">1.0</span></span>|
|[<span data-ttu-id="f8455-129">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-130">ReadItem</span></span>|
|[<span data-ttu-id="f8455-131">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-133">例</span><span class="sxs-lookup"><span data-stu-id="f8455-133">Example</span></span>

<span data-ttu-id="f8455-134">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="f8455-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="f8455-135">bcc:[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="f8455-136">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f8455-137">新規作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="f8455-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-138">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-138">Type:</span></span>

*   [<span data-ttu-id="f8455-139">受信者</span><span class="sxs-lookup"><span data-stu-id="f8455-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f8455-140">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-140">Requirements</span></span>

|<span data-ttu-id="f8455-141">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-141">Requirement</span></span>| <span data-ttu-id="f8455-142">値</span><span class="sxs-lookup"><span data-stu-id="f8455-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-143">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f8455-144">1.1</span></span>|
|[<span data-ttu-id="f8455-145">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-146">ReadItem</span></span>|
|[<span data-ttu-id="f8455-147">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-148">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-149">例</span><span class="sxs-lookup"><span data-stu-id="f8455-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="f8455-150">本文:[本文](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="f8455-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="f8455-151">アイテムの本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-152">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-152">Type:</span></span>

*   [<span data-ttu-id="f8455-153">本文</span><span class="sxs-lookup"><span data-stu-id="f8455-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="f8455-154">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-154">Requirements</span></span>

|<span data-ttu-id="f8455-155">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-155">Requirement</span></span>| <span data-ttu-id="f8455-156">値</span><span class="sxs-lookup"><span data-stu-id="f8455-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-157">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-158">1.1</span><span class="sxs-lookup"><span data-stu-id="f8455-158">1.1</span></span>|
|[<span data-ttu-id="f8455-159">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-160">ReadItem</span></span>|
|[<span data-ttu-id="f8455-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-162">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="f8455-163">cc: 配列。 <[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="f8455-164">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f8455-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f8455-165">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f8455-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-166">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-166">Read mode</span></span>

<span data-ttu-id="f8455-p106">`cc`プロパティは、メッセージの**CC**列にある各受信者一覧の`EmailAddressDetails`オブジェクトを含む配列を返します。コレクションは最大100個のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-169">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-169">Compose mode</span></span>

<span data-ttu-id="f8455-170">`cc` プロパティは、メッセージの **CC** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-171">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-171">Type:</span></span>

*   <span data-ttu-id="f8455-172">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-173">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-173">Requirements</span></span>

|<span data-ttu-id="f8455-174">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-174">Requirement</span></span>| <span data-ttu-id="f8455-175">値</span><span class="sxs-lookup"><span data-stu-id="f8455-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-177">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-177">1.0</span></span>|
|[<span data-ttu-id="f8455-178">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-179">ReadItem</span></span>|
|[<span data-ttu-id="f8455-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-181">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-182">例</span><span class="sxs-lookup"><span data-stu-id="f8455-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f8455-183">（空白が可能）conversationId：文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="f8455-184">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f8455-p107">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティで整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="f8455-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f8455-p108">作成フォームの新しいアイテムに対してこのプロパティの Null を取得します。ユーザーが件名を設定し項目を保存する場合、`conversationId`プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-189">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-189">Type:</span></span>

*   <span data-ttu-id="f8455-190">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-191">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-191">Requirements</span></span>

|<span data-ttu-id="f8455-192">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-192">Requirement</span></span>| <span data-ttu-id="f8455-193">値</span><span class="sxs-lookup"><span data-stu-id="f8455-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-194">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-195">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-195">1.0</span></span>|
|[<span data-ttu-id="f8455-196">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-197">ReadItem</span></span>|
|[<span data-ttu-id="f8455-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-199">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f8455-200">dateTimeCreated: 日付</span><span class="sxs-lookup"><span data-stu-id="f8455-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="f8455-p109">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-203">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-203">Type:</span></span>

*   <span data-ttu-id="f8455-204">日付</span><span class="sxs-lookup"><span data-stu-id="f8455-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-205">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-205">Requirements</span></span>

|<span data-ttu-id="f8455-206">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-206">Requirement</span></span>| <span data-ttu-id="f8455-207">値</span><span class="sxs-lookup"><span data-stu-id="f8455-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-209">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-209">1.0</span></span>|
|[<span data-ttu-id="f8455-210">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-211">ReadItem</span></span>|
|[<span data-ttu-id="f8455-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-213">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-214">例</span><span class="sxs-lookup"><span data-stu-id="f8455-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f8455-215">dateTimeModified: 日付</span><span class="sxs-lookup"><span data-stu-id="f8455-215">dateTimeModified :Date</span></span>

<span data-ttu-id="f8455-p110">アイテムが最後に変更された日時を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-218">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-219">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-219">Type:</span></span>

*   <span data-ttu-id="f8455-220">日付</span><span class="sxs-lookup"><span data-stu-id="f8455-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-221">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-221">Requirements</span></span>

|<span data-ttu-id="f8455-222">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-222">Requirement</span></span>| <span data-ttu-id="f8455-223">値</span><span class="sxs-lookup"><span data-stu-id="f8455-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-224">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-225">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-225">1.0</span></span>|
|[<span data-ttu-id="f8455-226">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-227">ReadItem</span></span>|
|[<span data-ttu-id="f8455-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-229">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-230">例</span><span class="sxs-lookup"><span data-stu-id="f8455-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="f8455-231">end:日付|[時間](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8455-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="f8455-232">予定が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f8455-p111">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-235">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-235">Read mode</span></span>

<span data-ttu-id="f8455-236">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-237">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-237">Compose mode</span></span>

<span data-ttu-id="f8455-238">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f8455-239">[ `Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8455-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-240">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-240">Type:</span></span>

*   <span data-ttu-id="f8455-241">日付| [時間](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8455-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-242">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-242">Requirements</span></span>

|<span data-ttu-id="f8455-243">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-243">Requirement</span></span>| <span data-ttu-id="f8455-244">値</span><span class="sxs-lookup"><span data-stu-id="f8455-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-245">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-246">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-246">1.0</span></span>|
|[<span data-ttu-id="f8455-247">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-248">ReadItem</span></span>|
|[<span data-ttu-id="f8455-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-250">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-251">例</span><span class="sxs-lookup"><span data-stu-id="f8455-251">Example</span></span>

<span data-ttu-id="f8455-252">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="f8455-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8455-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8455-p112">メッセージの送信者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f8455-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-258">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="f8455-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-259">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-259">Type:</span></span>

*   [<span data-ttu-id="f8455-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8455-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8455-261">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-261">Requirements</span></span>

|<span data-ttu-id="f8455-262">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-262">Requirement</span></span>| <span data-ttu-id="f8455-263">値</span><span class="sxs-lookup"><span data-stu-id="f8455-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-265">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-265">1.0</span></span>|
|[<span data-ttu-id="f8455-266">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-267">ReadItem</span></span>|
|[<span data-ttu-id="f8455-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-269">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f8455-270">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-270">internetMessageId :String</span></span>

<span data-ttu-id="f8455-p114">電子メール メッセージのインターネット メッセージ 識別子を取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-273">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-273">Type:</span></span>

*   <span data-ttu-id="f8455-274">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-275">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-275">Requirements</span></span>

|<span data-ttu-id="f8455-276">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-276">Requirement</span></span>| <span data-ttu-id="f8455-277">値</span><span class="sxs-lookup"><span data-stu-id="f8455-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-278">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-279">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-279">1.0</span></span>|
|[<span data-ttu-id="f8455-280">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-281">ReadItem</span></span>|
|[<span data-ttu-id="f8455-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-283">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-284">例</span><span class="sxs-lookup"><span data-stu-id="f8455-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f8455-285">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-285">itemClass :String</span></span>

<span data-ttu-id="f8455-p115">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f8455-p116">`itemClass` プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f8455-290">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-290">Type</span></span> | <span data-ttu-id="f8455-291">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-291">Description</span></span> | <span data-ttu-id="f8455-292">項目のクラス</span><span class="sxs-lookup"><span data-stu-id="f8455-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f8455-293">予定アイテム</span><span class="sxs-lookup"><span data-stu-id="f8455-293">Appointment items</span></span> | <span data-ttu-id="f8455-294">これらは、アイテム クラス `IPM.Appointment` または `IPM.Appointment.Occurence` の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="f8455-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="f8455-295">メッセージ アイテム</span><span class="sxs-lookup"><span data-stu-id="f8455-295">Message items</span></span> | <span data-ttu-id="f8455-296">これには、基本のメッセージ クラス として `IPM.Schedule.Meeting`  を使用する、既定のメッセージ クラス `IPM.Note`  会議出席依頼、返信および取り消しを持つ電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f8455-297">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` など)を作成できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-298">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-298">Type:</span></span>

*   <span data-ttu-id="f8455-299">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-300">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-300">Requirements</span></span>

|<span data-ttu-id="f8455-301">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-301">Requirement</span></span>| <span data-ttu-id="f8455-302">値</span><span class="sxs-lookup"><span data-stu-id="f8455-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-303">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-304">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-304">1.0</span></span>|
|[<span data-ttu-id="f8455-305">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-306">ReadItem</span></span>|
|[<span data-ttu-id="f8455-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-308">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-309">例</span><span class="sxs-lookup"><span data-stu-id="f8455-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f8455-310">（空白が可能） itemId ：文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-310">(nullable) itemId :String</span></span>

<span data-ttu-id="f8455-p117">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-313">`itemId` プロパティから返される識別子は、Exchange Web サービスのアイテム識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="f8455-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f8455-314">`itemId` プロパティは、Outlook の Entry ID または Outlook の REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="f8455-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f8455-315">この値を使用して REST API の呼び出しを行う前に [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8455-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f8455-316">詳細については、 [Outlook アドインから Outlook REST API の使用](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f8455-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f8455-p119">新規作成モードでは、`itemId` プロパティは使用できません。アイテム識別子が必要な場合、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してアイテムをストアに保存できます。そうすると、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメーターでアイテム識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-319">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-319">Type:</span></span>

*   <span data-ttu-id="f8455-320">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-321">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-321">Requirements</span></span>

|<span data-ttu-id="f8455-322">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-322">Requirement</span></span>| <span data-ttu-id="f8455-323">値</span><span class="sxs-lookup"><span data-stu-id="f8455-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-324">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-325">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-325">1.0</span></span>|
|[<span data-ttu-id="f8455-326">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-327">ReadItem</span></span>|
|[<span data-ttu-id="f8455-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-329">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-330">例</span><span class="sxs-lookup"><span data-stu-id="f8455-330">Example</span></span>

<span data-ttu-id="f8455-p120">次のコードは、アイテム識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合、アイテムはストアに保存され、非同期の結果からアイテム識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="f8455-333">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f8455-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f8455-334">インスタンスが表しているアイテムの種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f8455-335">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="f8455-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-336">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-336">Type:</span></span>

*   [<span data-ttu-id="f8455-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f8455-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f8455-338">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-338">Requirements</span></span>

|<span data-ttu-id="f8455-339">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-339">Requirement</span></span>| <span data-ttu-id="f8455-340">値</span><span class="sxs-lookup"><span data-stu-id="f8455-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-341">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-342">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-342">1.0</span></span>|
|[<span data-ttu-id="f8455-343">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-344">ReadItem</span></span>|
|[<span data-ttu-id="f8455-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-346">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-347">例</span><span class="sxs-lookup"><span data-stu-id="f8455-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="f8455-348">位置: 文字列|[](/javascript/api/outlook_1_3/office.location)位置</span><span class="sxs-lookup"><span data-stu-id="f8455-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="f8455-349">予定の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-350">Read mode</span></span>

<span data-ttu-id="f8455-351">`location` プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-352">Compose mode</span></span>

<span data-ttu-id="f8455-353">`location` プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-354">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-354">Type:</span></span>

*   <span data-ttu-id="f8455-355">文字列 | [場所](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="f8455-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-356">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-356">Requirements</span></span>

|<span data-ttu-id="f8455-357">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-357">Requirement</span></span>| <span data-ttu-id="f8455-358">値</span><span class="sxs-lookup"><span data-stu-id="f8455-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-359">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-360">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-360">1.0</span></span>|
|[<span data-ttu-id="f8455-361">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-362">ReadItem</span></span>|
|[<span data-ttu-id="f8455-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-364">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-365">例</span><span class="sxs-lookup"><span data-stu-id="f8455-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f8455-366">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-366">normalizedSubject :String</span></span>

<span data-ttu-id="f8455-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f8455-p122">normalizedSubject プロパティは、電子メール プログラムによって標準のプレフィックス (`RE:` や `FW:` など) が追加されたアイテムの件名を取得します。プレフィックスが付いたままの状態でアイテムの件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-371">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-371">Type:</span></span>

*   <span data-ttu-id="f8455-372">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-373">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-373">Requirements</span></span>

|<span data-ttu-id="f8455-374">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-374">Requirement</span></span>| <span data-ttu-id="f8455-375">値</span><span class="sxs-lookup"><span data-stu-id="f8455-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-376">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-377">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-377">1.0</span></span>|
|[<span data-ttu-id="f8455-378">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-379">ReadItem</span></span>|
|[<span data-ttu-id="f8455-380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-381">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-382">例</span><span class="sxs-lookup"><span data-stu-id="f8455-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="f8455-383">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f8455-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="f8455-384">項目の通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-385">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-385">Type:</span></span>

*   [<span data-ttu-id="f8455-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f8455-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f8455-387">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-387">Requirements</span></span>

|<span data-ttu-id="f8455-388">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-388">Requirement</span></span>| <span data-ttu-id="f8455-389">値</span><span class="sxs-lookup"><span data-stu-id="f8455-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-390">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-391">1.3</span><span class="sxs-lookup"><span data-stu-id="f8455-391">1.3</span></span>|
|[<span data-ttu-id="f8455-392">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-393">ReadItem</span></span>|
|[<span data-ttu-id="f8455-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-395">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="f8455-396">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="f8455-397">イベントの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f8455-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f8455-398">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f8455-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-399">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-399">Read mode</span></span>

<span data-ttu-id="f8455-400">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-401">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-401">Compose mode</span></span>

<span data-ttu-id="f8455-402">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-403">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-403">Type:</span></span>

*   <span data-ttu-id="f8455-404">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-405">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-405">Requirements</span></span>

|<span data-ttu-id="f8455-406">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-406">Requirement</span></span>| <span data-ttu-id="f8455-407">値</span><span class="sxs-lookup"><span data-stu-id="f8455-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-408">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-409">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-409">1.0</span></span>|
|[<span data-ttu-id="f8455-410">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-411">ReadItem</span></span>|
|[<span data-ttu-id="f8455-412">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-413">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-414">例</span><span class="sxs-lookup"><span data-stu-id="f8455-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="f8455-415">開催者:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8455-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8455-p124">指定の会議の開催者の電子メール アドレスを取得します。読み取り専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-418">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-418">Type:</span></span>

*   [<span data-ttu-id="f8455-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8455-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8455-420">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-420">Requirements</span></span>

|<span data-ttu-id="f8455-421">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-421">Requirement</span></span>| <span data-ttu-id="f8455-422">値</span><span class="sxs-lookup"><span data-stu-id="f8455-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-423">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-424">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-424">1.0</span></span>|
|[<span data-ttu-id="f8455-425">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-426">ReadItem</span></span>|
|[<span data-ttu-id="f8455-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-428">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-429">例</span><span class="sxs-lookup"><span data-stu-id="f8455-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="f8455-430">requiredAttendees: 配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="f8455-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="f8455-431">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f8455-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f8455-432">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f8455-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-433">Read mode</span></span>

<span data-ttu-id="f8455-434">`requiredAttendees` プロパティは、会議への各必須出席者の `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-435">Compose mode</span></span>

<span data-ttu-id="f8455-436">`requiredAttendees` プロパティは会議への必須出席者を取得または設定するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-437">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-437">Type:</span></span>

*   <span data-ttu-id="f8455-438">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-439">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-439">Requirements</span></span>

|<span data-ttu-id="f8455-440">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-440">Requirement</span></span>| <span data-ttu-id="f8455-441">値</span><span class="sxs-lookup"><span data-stu-id="f8455-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-442">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-443">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-443">1.0</span></span>|
|[<span data-ttu-id="f8455-444">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-445">ReadItem</span></span>|
|[<span data-ttu-id="f8455-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-447">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-448">例</span><span class="sxs-lookup"><span data-stu-id="f8455-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="f8455-449">送信者:[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f8455-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="f8455-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f8455-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) プロパティと `sender` プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、sender プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-454">`sender`プロパティ内の `EmailAddressDetails` オブジェクトの `recipientType`プロパティは、`undefined`です。</span><span class="sxs-lookup"><span data-stu-id="f8455-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-455">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-455">Type:</span></span>

*   [<span data-ttu-id="f8455-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f8455-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f8455-457">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-457">Requirements</span></span>

|<span data-ttu-id="f8455-458">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-458">Requirement</span></span>| <span data-ttu-id="f8455-459">値</span><span class="sxs-lookup"><span data-stu-id="f8455-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-460">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-461">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-461">1.0</span></span>|
|[<span data-ttu-id="f8455-462">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-463">ReadItem</span></span>|
|[<span data-ttu-id="f8455-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-465">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-466">例</span><span class="sxs-lookup"><span data-stu-id="f8455-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="f8455-467">開始: 日付 |[ 時間](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8455-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="f8455-468">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f8455-p128">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-471">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-471">Read mode</span></span>

<span data-ttu-id="f8455-472">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-473">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-473">Compose mode</span></span>

<span data-ttu-id="f8455-474">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f8455-475">[ `Time.setAsync` ](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f8455-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-476">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-476">Type:</span></span>

*   <span data-ttu-id="f8455-477">日付| [時間](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="f8455-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-478">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-478">Requirements</span></span>

|<span data-ttu-id="f8455-479">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-479">Requirement</span></span>| <span data-ttu-id="f8455-480">値</span><span class="sxs-lookup"><span data-stu-id="f8455-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-481">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-482">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-482">1.0</span></span>|
|[<span data-ttu-id="f8455-483">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-484">ReadItem</span></span>|
|[<span data-ttu-id="f8455-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-486">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-487">例</span><span class="sxs-lookup"><span data-stu-id="f8455-487">Example</span></span>

<span data-ttu-id="f8455-488">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="f8455-489">件名: 文字列 | [件名](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f8455-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="f8455-490">アイテムの件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f8455-491">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="f8455-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-492">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-492">Read mode</span></span>

<span data-ttu-id="f8455-p129">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f8455-495">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-495">Compose mode</span></span>

<span data-ttu-id="f8455-496">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f8455-497">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-497">Type:</span></span>

*   <span data-ttu-id="f8455-498">文字列 | [件名](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f8455-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-499">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-499">Requirements</span></span>

|<span data-ttu-id="f8455-500">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-500">Requirement</span></span>| <span data-ttu-id="f8455-501">値</span><span class="sxs-lookup"><span data-stu-id="f8455-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-502">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-503">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-503">1.0</span></span>|
|[<span data-ttu-id="f8455-504">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-505">ReadItem</span></span>|
|[<span data-ttu-id="f8455-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-507">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="f8455-508">to: 配列。 [EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="f8455-509">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f8455-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f8455-510">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="f8455-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f8455-511">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="f8455-511">Read mode</span></span>

<span data-ttu-id="f8455-p131">`to` プロパティは、メッセージの **To** 行にある各受信者について、`EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは最大 100 メンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f8455-514">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="f8455-514">Compose mode</span></span>

<span data-ttu-id="f8455-515">`to` プロパティは、メッセージの **To** 行にある受信者を取得または更新するメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f8455-516">型:</span><span class="sxs-lookup"><span data-stu-id="f8455-516">Type:</span></span>

*   <span data-ttu-id="f8455-517">配列 。<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> |[受信者](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f8455-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-518">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-518">Requirements</span></span>

|<span data-ttu-id="f8455-519">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-519">Requirement</span></span>| <span data-ttu-id="f8455-520">値</span><span class="sxs-lookup"><span data-stu-id="f8455-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-521">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-522">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-522">1.0</span></span>|
|[<span data-ttu-id="f8455-523">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-524">ReadItem</span></span>|
|[<span data-ttu-id="f8455-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-526">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-527">例</span><span class="sxs-lookup"><span data-stu-id="f8455-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f8455-528">メソッド</span><span class="sxs-lookup"><span data-stu-id="f8455-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f8455-529">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="f8455-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f8455-530">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f8455-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f8455-531">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="f8455-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f8455-532">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-533">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-533">Parameters:</span></span>

|<span data-ttu-id="f8455-534">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-534">Name</span></span>| <span data-ttu-id="f8455-535">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-535">Type</span></span>| <span data-ttu-id="f8455-536">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-536">Attributes</span></span>| <span data-ttu-id="f8455-537">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f8455-538">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-538">String</span></span>||<span data-ttu-id="f8455-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="f8455-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f8455-541">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-541">String</span></span>||<span data-ttu-id="f8455-p133">アップロード中に表示される添付ファイルがそのファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="f8455-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f8455-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-544">Object</span></span>| <span data-ttu-id="f8455-545">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-545">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-546">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-547">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-547">Object</span></span>| <span data-ttu-id="f8455-548">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-548">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-549">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8455-550">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-550">function</span></span>| <span data-ttu-id="f8455-551">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-551">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-552">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8455-553">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f8455-554">添付ファイルのアップロードに失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8455-555">エラー</span><span class="sxs-lookup"><span data-stu-id="f8455-555">Errors</span></span>

| <span data-ttu-id="f8455-556">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f8455-556">Error code</span></span> | <span data-ttu-id="f8455-557">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f8455-558">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="f8455-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f8455-559">許可されていない拡張子付きの添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f8455-560">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f8455-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-561">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-561">Requirements</span></span>

|<span data-ttu-id="f8455-562">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-562">Requirement</span></span>| <span data-ttu-id="f8455-563">値</span><span class="sxs-lookup"><span data-stu-id="f8455-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-564">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-565">1.1</span><span class="sxs-lookup"><span data-stu-id="f8455-565">1.1</span></span>|
|[<span data-ttu-id="f8455-566">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-569">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-570">例</span><span class="sxs-lookup"><span data-stu-id="f8455-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f8455-571">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="f8455-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f8455-572">メッセージなどの Exchange アイテムを添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="f8455-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f8455-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` というパラメータがあるメソッドが呼び出されます。このパラメーターには、添付ファイルの識別子、または項目を添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f8455-576">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f8455-577">Office アドインを Outlook Web App で実行している場合、`addItemAttachmentAsync` メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-578">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-578">Parameters:</span></span>

|<span data-ttu-id="f8455-579">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-579">Name</span></span>| <span data-ttu-id="f8455-580">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-580">Type</span></span>| <span data-ttu-id="f8455-581">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-581">Attributes</span></span>| <span data-ttu-id="f8455-582">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f8455-583">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-583">String</span></span>||<span data-ttu-id="f8455-p135">添付するアイテムの Exchange 識別子です。最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f8455-586">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-586">String</span></span>||<span data-ttu-id="f8455-p136">添付するアイテムの件名です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f8455-589">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-589">Object</span></span>| <span data-ttu-id="f8455-590">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-590">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-591">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-592">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-592">Object</span></span>| <span data-ttu-id="f8455-593">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-593">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-594">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8455-595">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-595">function</span></span>| <span data-ttu-id="f8455-596">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-596">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-597">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8455-598">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティに設定されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f8455-599">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult` オブジェクトが `Error` オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8455-600">エラー</span><span class="sxs-lookup"><span data-stu-id="f8455-600">Errors</span></span>

| <span data-ttu-id="f8455-601">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f8455-601">Error code</span></span> | <span data-ttu-id="f8455-602">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f8455-603">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="f8455-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-604">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-604">Requirements</span></span>

|<span data-ttu-id="f8455-605">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-605">Requirement</span></span>| <span data-ttu-id="f8455-606">値</span><span class="sxs-lookup"><span data-stu-id="f8455-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-607">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-608">1.1</span><span class="sxs-lookup"><span data-stu-id="f8455-608">1.1</span></span>|
|[<span data-ttu-id="f8455-609">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-612">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-613">例</span><span class="sxs-lookup"><span data-stu-id="f8455-613">Example</span></span>

<span data-ttu-id="f8455-614">次の例では、既存の Outlook アイテムが名前 `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="f8455-615">閉じる()</span><span class="sxs-lookup"><span data-stu-id="f8455-615">close()</span></span>

<span data-ttu-id="f8455-616">新規作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="f8455-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f8455-p137">`close`メソッドの動作は、作成中のアイテムの現在の状態によって異なります。アイテムに未保存の変更がある場合は、クライアントはユーザーに対して閉じる操作を保存、破棄、またはキャンセルするように求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-619">Outlook on the webでは、項目が予定で、`saveAsync`を用いて事前に保存されている場合、項目が最後に保存されてから何も変更されていない場合でも、ユーザーに対して保存、破棄またはキャンセルするよう求めるプロンプトを表示します。</span><span class="sxs-lookup"><span data-stu-id="f8455-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f8455-620">Outlook デスクトップ クライアントでは、メッセージがインライン返信の場合、`close`メソッドは無効になります。</span><span class="sxs-lookup"><span data-stu-id="f8455-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-621">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-621">Requirements</span></span>

|<span data-ttu-id="f8455-622">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-622">Requirement</span></span>| <span data-ttu-id="f8455-623">値</span><span class="sxs-lookup"><span data-stu-id="f8455-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-624">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-625">1.3</span><span class="sxs-lookup"><span data-stu-id="f8455-625">1.3</span></span>|
|[<span data-ttu-id="f8455-626">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-627">制限あり</span><span class="sxs-lookup"><span data-stu-id="f8455-627">Restricted</span></span>|
|[<span data-ttu-id="f8455-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-629">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f8455-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f8455-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f8455-631">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した回答フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-632">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f8455-633">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f8455-634">文字列パラメーターのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f8455-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f8455-p138">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f8455-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-638">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-638">Parameters:</span></span>

|<span data-ttu-id="f8455-639">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-639">Name</span></span>| <span data-ttu-id="f8455-640">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-640">Type</span></span>| <span data-ttu-id="f8455-641">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f8455-642">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-642">String &#124; Object</span></span>| |<span data-ttu-id="f8455-p139">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f8455-645">**または**</span><span class="sxs-lookup"><span data-stu-id="f8455-645">**OR**</span></span><br/><span data-ttu-id="f8455-p140">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f8455-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f8455-648">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-648">String</span></span> | <span data-ttu-id="f8455-649">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-649">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-p141">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f8455-652">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f8455-653">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-653">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-654">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f8455-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f8455-655">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-655">String</span></span> | | <span data-ttu-id="f8455-p142">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="f8455-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f8455-658">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-658">String</span></span> | | <span data-ttu-id="f8455-659">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f8455-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f8455-660">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-660">String</span></span> | | <span data-ttu-id="f8455-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="f8455-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f8455-663">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-663">String</span></span> | | <span data-ttu-id="f8455-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f8455-667">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-667">function</span></span> | <span data-ttu-id="f8455-668">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-668">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-669">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-670">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-670">Requirements</span></span>

|<span data-ttu-id="f8455-671">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-671">Requirement</span></span>| <span data-ttu-id="f8455-672">値</span><span class="sxs-lookup"><span data-stu-id="f8455-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-673">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-674">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-674">1.0</span></span>|
|[<span data-ttu-id="f8455-675">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-676">ReadItem</span></span>|
|[<span data-ttu-id="f8455-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-678">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8455-679">例</span><span class="sxs-lookup"><span data-stu-id="f8455-679">Examples</span></span>

<span data-ttu-id="f8455-680">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f8455-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f8455-681">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f8455-682">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f8455-683">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-683">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="f8455-684">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-684">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="f8455-685">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f8455-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f8455-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="f8455-687">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-688">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f8455-689">Outlook Web App では、回答フォームは、3 列表示のポップアウト形式、および 2 列または 1 列表示のポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f8455-690">文字列パラメーターのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="f8455-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f8455-p145">`formData.attachments` パラメーターで添付ファイルを指定すると、Outlook と Office Web Apps はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="f8455-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-694">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-694">Parameters:</span></span>

|<span data-ttu-id="f8455-695">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-695">Name</span></span>| <span data-ttu-id="f8455-696">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-696">Type</span></span>| <span data-ttu-id="f8455-697">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f8455-698">文字列 |オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-698">String &#124; Object</span></span>| | <span data-ttu-id="f8455-p146">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f8455-701">**または**</span><span class="sxs-lookup"><span data-stu-id="f8455-701">**OR**</span></span><br/><span data-ttu-id="f8455-p147">本文または添付ファイルのデータと、コールバック関数を格納しているオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="f8455-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f8455-704">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-704">String</span></span> | <span data-ttu-id="f8455-705">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-705">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-p148">回答フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="f8455-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f8455-708">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f8455-709">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-709">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-710">ファイルまたはアイテムの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="f8455-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f8455-711">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-711">String</span></span> | | <span data-ttu-id="f8455-p149">添付ファイルの種類を示します。ファイルの添付ファイルの場合は `file`、アイテムの添付ファイルの場合は `item` でなければいけません。</span><span class="sxs-lookup"><span data-stu-id="f8455-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f8455-714">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-714">String</span></span> | | <span data-ttu-id="f8455-715">添付ファイル名を含む文字列です。最大長は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="f8455-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f8455-716">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-716">String</span></span> | | <span data-ttu-id="f8455-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="f8455-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f8455-719">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-719">String</span></span> | | <span data-ttu-id="f8455-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大長が 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f8455-723">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-723">function</span></span> | <span data-ttu-id="f8455-724">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-724">&lt;optional&gt;</span></span> | <span data-ttu-id="f8455-725">メソッドが完了すると、`callback` パラメーターに渡された関数が、[AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトである シングル パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-726">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-726">Requirements</span></span>

|<span data-ttu-id="f8455-727">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-727">Requirement</span></span>| <span data-ttu-id="f8455-728">値</span><span class="sxs-lookup"><span data-stu-id="f8455-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-729">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-730">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-730">1.0</span></span>|
|[<span data-ttu-id="f8455-731">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-732">ReadItem</span></span>|
|[<span data-ttu-id="f8455-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-734">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8455-735">例</span><span class="sxs-lookup"><span data-stu-id="f8455-735">Examples</span></span>

<span data-ttu-id="f8455-736">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="f8455-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f8455-737">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f8455-738">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f8455-739">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-739">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="f8455-740">本文とアイテムの添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-740">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="f8455-741">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="f8455-742">getEntities() → {[エンティティ](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f8455-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="f8455-743">選択したアイテムの本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-744">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-745">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-745">Requirements</span></span>

|<span data-ttu-id="f8455-746">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-746">Requirement</span></span>| <span data-ttu-id="f8455-747">値</span><span class="sxs-lookup"><span data-stu-id="f8455-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-748">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-749">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-749">1.0</span></span>|
|[<span data-ttu-id="f8455-750">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-751">ReadItem</span></span>|
|[<span data-ttu-id="f8455-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-753">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-754">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-754">Returns:</span></span>

<span data-ttu-id="f8455-755">種類: [エンティティ](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f8455-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f8455-756">例</span><span class="sxs-lookup"><span data-stu-id="f8455-756">Example</span></span>

<span data-ttu-id="f8455-757">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="f8455-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="f8455-758">getEntitiesByType(entityType)] → [(空白可能) {<(String|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="f8455-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f8455-759">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-760">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-761">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-761">Parameters:</span></span>

|<span data-ttu-id="f8455-762">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-762">Name</span></span>| <span data-ttu-id="f8455-763">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-763">Type</span></span>| <span data-ttu-id="f8455-764">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f8455-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f8455-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="f8455-766">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="f8455-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-767">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-767">Requirements</span></span>

|<span data-ttu-id="f8455-768">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-768">Requirement</span></span>| <span data-ttu-id="f8455-769">値</span><span class="sxs-lookup"><span data-stu-id="f8455-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-770">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-771">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-771">1.0</span></span>|
|[<span data-ttu-id="f8455-772">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-773">制限あり</span><span class="sxs-lookup"><span data-stu-id="f8455-773">Restricted</span></span>|
|[<span data-ttu-id="f8455-774">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-775">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-776">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-776">Returns:</span></span>

<span data-ttu-id="f8455-777">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f8455-778">指定した型のエンティティがアイテムの本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="f8455-779">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="f8455-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f8455-780">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="f8455-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f8455-781">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="f8455-781">Value of `entityType`</span></span> | <span data-ttu-id="f8455-782">返される配列内のオブジェクトの型</span><span class="sxs-lookup"><span data-stu-id="f8455-782">Type of objects in returned array</span></span> | <span data-ttu-id="f8455-783">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f8455-784">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-784">String</span></span> | <span data-ttu-id="f8455-785">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f8455-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f8455-786">連絡先</span><span class="sxs-lookup"><span data-stu-id="f8455-786">Contact</span></span> | <span data-ttu-id="f8455-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8455-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f8455-788">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-788">String</span></span> | <span data-ttu-id="f8455-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8455-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f8455-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f8455-790">MeetingSuggestion</span></span> | <span data-ttu-id="f8455-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8455-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f8455-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f8455-792">PhoneNumber</span></span> | <span data-ttu-id="f8455-793">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f8455-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f8455-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f8455-794">TaskSuggestion</span></span> | <span data-ttu-id="f8455-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f8455-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f8455-796">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-796">String</span></span> | <span data-ttu-id="f8455-797">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="f8455-797">**Restricted**</span></span> |

<span data-ttu-id="f8455-798">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f8455-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f8455-799">例</span><span class="sxs-lookup"><span data-stu-id="f8455-799">Example</span></span>

<span data-ttu-id="f8455-800">次の例は、現在のアイテムの本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="f8455-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="f8455-801">getFilteredEntitiesByName(name)] → [(Null 許容) {<(文字列| [連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f8455-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f8455-802">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-803">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f8455-804">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-805">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-805">Parameters:</span></span>

|<span data-ttu-id="f8455-806">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-806">Name</span></span>| <span data-ttu-id="f8455-807">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-807">Type</span></span>| <span data-ttu-id="f8455-808">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f8455-809">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-809">String</span></span>|<span data-ttu-id="f8455-810">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="f8455-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-811">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-811">Requirements</span></span>

|<span data-ttu-id="f8455-812">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-812">Requirement</span></span>| <span data-ttu-id="f8455-813">値</span><span class="sxs-lookup"><span data-stu-id="f8455-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-814">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-815">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-815">1.0</span></span>|
|[<span data-ttu-id="f8455-816">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-817">ReadItem</span></span>|
|[<span data-ttu-id="f8455-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-819">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-820">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-820">Returns:</span></span>

<span data-ttu-id="f8455-p153">`FilterName` 要素の値が `name` パラメーターと一致するマニフェスト内に `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致せず、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f8455-823">型:Array.<(文字列|[連絡先](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f8455-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f8455-824">getRegExMatches() → {オブジェクト。</span><span class="sxs-lookup"><span data-stu-id="f8455-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f8455-825">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-826">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f8455-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f8455-830">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="f8455-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f8455-831">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティがありえます。</span><span class="sxs-lookup"><span data-stu-id="f8455-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f8455-p155">アイテムの 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、アイテムの本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8455-835">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-835">Requirements</span></span>

|<span data-ttu-id="f8455-836">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-836">Requirement</span></span>| <span data-ttu-id="f8455-837">値</span><span class="sxs-lookup"><span data-stu-id="f8455-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-838">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-839">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-839">1.0</span></span>|
|[<span data-ttu-id="f8455-840">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-841">ReadItem</span></span>|
|[<span data-ttu-id="f8455-842">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-843">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-844">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-844">Returns:</span></span>

<span data-ttu-id="f8455-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="f8455-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f8455-847">

<dt>型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f8455-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8455-848">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8455-849">例</span><span class="sxs-lookup"><span data-stu-id="f8455-849">Example</span></span>

<span data-ttu-id="f8455-850">次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="f8455-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f8455-851">getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="f8455-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f8455-852">選択したアイテム内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-853">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f8455-854">`getRegExMatchesByName` メソッドは、`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。このルール要素は、指定された `RegExName` 要素値を持つマニフェスト XML ファイル内にあります。</span><span class="sxs-lookup"><span data-stu-id="f8455-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f8455-p157">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="f8455-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-857">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-857">Parameters:</span></span>

|<span data-ttu-id="f8455-858">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-858">Name</span></span>| <span data-ttu-id="f8455-859">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-859">Type</span></span>| <span data-ttu-id="f8455-860">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f8455-861">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-861">String</span></span>|<span data-ttu-id="f8455-862">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="f8455-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-863">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-863">Requirements</span></span>

|<span data-ttu-id="f8455-864">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-864">Requirement</span></span>| <span data-ttu-id="f8455-865">値</span><span class="sxs-lookup"><span data-stu-id="f8455-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-866">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-867">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-867">1.0</span></span>|
|[<span data-ttu-id="f8455-868">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-869">ReadItem</span></span>|
|[<span data-ttu-id="f8455-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-871">読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-872">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-872">Returns:</span></span>

<span data-ttu-id="f8455-873">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="f8455-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f8455-874">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="f8455-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8455-875">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="f8455-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8455-876">例</span><span class="sxs-lookup"><span data-stu-id="f8455-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f8455-877">getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="f8455-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f8455-878">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f8455-p158">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-881">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-881">Parameters:</span></span>

|<span data-ttu-id="f8455-882">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-882">Name</span></span>| <span data-ttu-id="f8455-883">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-883">Type</span></span>| <span data-ttu-id="f8455-884">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-884">Attributes</span></span>| <span data-ttu-id="f8455-885">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f8455-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f8455-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f8455-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f8455-890">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-890">Object</span></span>| <span data-ttu-id="f8455-891">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-891">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-892">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-893">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-893">Object</span></span>| <span data-ttu-id="f8455-894">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-894">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-895">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8455-896">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-896">function</span></span>||<span data-ttu-id="f8455-897">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8455-898">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="f8455-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f8455-899">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは <   または `body` `subject` になります。</span><span class="sxs-lookup"><span data-stu-id="f8455-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-900">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-900">Requirements</span></span>

|<span data-ttu-id="f8455-901">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-901">Requirement</span></span>| <span data-ttu-id="f8455-902">値</span><span class="sxs-lookup"><span data-stu-id="f8455-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-903">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-904">1.2</span><span class="sxs-lookup"><span data-stu-id="f8455-904">1.2</span></span>|
|[<span data-ttu-id="f8455-905">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-908">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f8455-909">戻り値 :</span><span class="sxs-lookup"><span data-stu-id="f8455-909">Returns:</span></span>

<span data-ttu-id="f8455-910">`coercionType`で決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="f8455-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f8455-911">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="f8455-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f8455-912">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f8455-913">例</span><span class="sxs-lookup"><span data-stu-id="f8455-913">Example</span></span>

```
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f8455-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f8455-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f8455-915">選択されたアイテムのこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="f8455-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f8455-p161">カスタム プロパティは、アプリケーションごと、アイテムごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="f8455-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-919">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-919">Parameters:</span></span>

|<span data-ttu-id="f8455-920">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-920">Name</span></span>| <span data-ttu-id="f8455-921">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-921">Type</span></span>| <span data-ttu-id="f8455-922">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-922">Attributes</span></span>| <span data-ttu-id="f8455-923">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f8455-924">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-924">function</span></span>||<span data-ttu-id="f8455-925">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8455-926">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f8455-927">項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトが使用できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f8455-928">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-928">Object</span></span>| <span data-ttu-id="f8455-929">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-929">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-930">開発者は、コールバック 関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="f8455-931">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="f8455-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-932">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-932">Requirements</span></span>

|<span data-ttu-id="f8455-933">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-933">Requirement</span></span>| <span data-ttu-id="f8455-934">値</span><span class="sxs-lookup"><span data-stu-id="f8455-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-935">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-936">1.0以降</span><span class="sxs-lookup"><span data-stu-id="f8455-936">1.0</span></span>|
|[<span data-ttu-id="f8455-937">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f8455-938">ReadItem</span></span>|
|[<span data-ttu-id="f8455-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-940">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="f8455-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-941">例</span><span class="sxs-lookup"><span data-stu-id="f8455-941">Example</span></span>

<span data-ttu-id="f8455-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f8455-945">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="f8455-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f8455-946">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f8455-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f8455-p165">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="f8455-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-951">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-951">Parameters:</span></span>

|<span data-ttu-id="f8455-952">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-952">Name</span></span>| <span data-ttu-id="f8455-953">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-953">Type</span></span>| <span data-ttu-id="f8455-954">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-954">Attributes</span></span>| <span data-ttu-id="f8455-955">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f8455-956">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-956">String</span></span>||<span data-ttu-id="f8455-p166">削除する添付ファイルの識別子です。文字列の最大長は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="f8455-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="f8455-959">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-959">Object</span></span>| <span data-ttu-id="f8455-960">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-960">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-961">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-962">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-962">Object</span></span>| <span data-ttu-id="f8455-963">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-963">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-964">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8455-965">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-965">function</span></span>| <span data-ttu-id="f8455-966">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-966">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-967">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f8455-968">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f8455-969">エラー</span><span class="sxs-lookup"><span data-stu-id="f8455-969">Errors</span></span>

| <span data-ttu-id="f8455-970">エラー コード</span><span class="sxs-lookup"><span data-stu-id="f8455-970">Error code</span></span> | <span data-ttu-id="f8455-971">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f8455-972">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="f8455-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-973">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-973">Requirements</span></span>

|<span data-ttu-id="f8455-974">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-974">Requirement</span></span>| <span data-ttu-id="f8455-975">値</span><span class="sxs-lookup"><span data-stu-id="f8455-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-976">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-977">1.1</span><span class="sxs-lookup"><span data-stu-id="f8455-977">1.1</span></span>|
|[<span data-ttu-id="f8455-978">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-980">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-981">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-982">例</span><span class="sxs-lookup"><span data-stu-id="f8455-982">Example</span></span>

<span data-ttu-id="f8455-983">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="f8455-983">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f8455-984">saveAsync ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="f8455-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="f8455-985">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="f8455-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="f8455-p167">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由でアイテム ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーにアイテムが保存されます。キャッシュ モードの Outlook では、ローカル キャッシュにアイテムが保存されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-989">アドインが、WS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、アイテムがサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f8455-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="f8455-990">項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f8455-p169">予定はドラフト状態にはならないため、作成モードで予定に`saveAsync`が呼び出される場合、そのアイテムはユーザーの予定表に通常の予定として保存されます。以前に保存されていない新しい予定の場合、招待状は送信されません。既存の予定を保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f8455-994">次のクライアントは、新規作成モードで予定上の `saveAsync` に対して様々なふるまいをします。</span><span class="sxs-lookup"><span data-stu-id="f8455-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f8455-995">Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="f8455-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="f8455-996">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f8455-997">作成モードの予定上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="f8455-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-998">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-998">Parameters:</span></span>

|<span data-ttu-id="f8455-999">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-999">Name</span></span>| <span data-ttu-id="f8455-1000">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-1000">Type</span></span>| <span data-ttu-id="f8455-1001">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-1001">Attributes</span></span>| <span data-ttu-id="f8455-1002">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f8455-1003">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-1003">Object</span></span>| <span data-ttu-id="f8455-1004">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-1005">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-1006">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-1006">Object</span></span>| <span data-ttu-id="f8455-1007">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-1008">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f8455-1009">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-1009">function</span></span>||<span data-ttu-id="f8455-1010">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f8455-1011">成功すると、アイテム識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8455-1012">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-1012">Requirements</span></span>

|<span data-ttu-id="f8455-1013">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-1013">Requirement</span></span>| <span data-ttu-id="f8455-1014">値</span><span class="sxs-lookup"><span data-stu-id="f8455-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-1015">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="f8455-1016">1.3</span></span>|
|[<span data-ttu-id="f8455-1017">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-1019">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-1020">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f8455-1021">例</span><span class="sxs-lookup"><span data-stu-id="f8455-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f8455-p171">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、アイテムのアイテム ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f8455-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f8455-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f8455-1025">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="f8455-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f8455-p172">`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f8455-1029">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="f8455-1029">Parameters:</span></span>

|<span data-ttu-id="f8455-1030">名前</span><span class="sxs-lookup"><span data-stu-id="f8455-1030">Name</span></span>| <span data-ttu-id="f8455-1031">種類</span><span class="sxs-lookup"><span data-stu-id="f8455-1031">Type</span></span>| <span data-ttu-id="f8455-1032">属性</span><span class="sxs-lookup"><span data-stu-id="f8455-1032">Attributes</span></span>| <span data-ttu-id="f8455-1033">説明</span><span class="sxs-lookup"><span data-stu-id="f8455-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f8455-1034">文字列</span><span class="sxs-lookup"><span data-stu-id="f8455-1034">String</span></span>||<span data-ttu-id="f8455-p173">挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f8455-1038">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-1038">Object</span></span>| <span data-ttu-id="f8455-1039">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-1040">次のプロパティのうち 1 つ以上を含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="f8455-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f8455-1041">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="f8455-1041">Object</span></span>| <span data-ttu-id="f8455-1042">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-1043">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="f8455-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="f8455-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f8455-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="f8455-1045">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="f8455-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="f8455-p174">`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f8455-p175">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f8455-1050">`coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="f8455-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f8455-1051">関数</span><span class="sxs-lookup"><span data-stu-id="f8455-1051">function</span></span>||<span data-ttu-id="f8455-1052">メソッドが完了すると、`callback`パラメータに渡された関数が、シングルパラメータ、`asyncResult`で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="f8455-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8455-1053">要件</span><span class="sxs-lookup"><span data-stu-id="f8455-1053">Requirements</span></span>

|<span data-ttu-id="f8455-1054">必要条件</span><span class="sxs-lookup"><span data-stu-id="f8455-1054">Requirement</span></span>| <span data-ttu-id="f8455-1055">値</span><span class="sxs-lookup"><span data-stu-id="f8455-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8455-1056">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f8455-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f8455-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="f8455-1057">1.2</span></span>|
|[<span data-ttu-id="f8455-1058">アクセス許可に必要なレベル</span><span class="sxs-lookup"><span data-stu-id="f8455-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f8455-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f8455-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="f8455-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f8455-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f8455-1061">新規作成</span><span class="sxs-lookup"><span data-stu-id="f8455-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f8455-1062">例</span><span class="sxs-lookup"><span data-stu-id="f8455-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```