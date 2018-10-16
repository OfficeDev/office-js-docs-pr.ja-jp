
# <a name="item"></a><span data-ttu-id="fa25d-101">項目</span><span class="sxs-lookup"><span data-stu-id="fa25d-101">item</span></span>

### <span data-ttu-id="fa25d-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fa25d-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="fa25d-p102">`item` 名前空間を使用して、現在選択されているメッセージ、会議出席依頼、または予約にアクセスします。 `item`の種類を[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype)プロパティを使用して決定できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-106">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-106">Requirements</span></span>

|<span data-ttu-id="fa25d-107">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-107">Requirement</span></span>| <span data-ttu-id="fa25d-108">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-109">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-109">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-110">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-110">1.0</span></span>|
|[<span data-ttu-id="fa25d-111">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-112">制限あり</span><span class="sxs-lookup"><span data-stu-id="fa25d-112">Restricted</span></span>|
|[<span data-ttu-id="fa25d-113">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-114">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="fa25d-115">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-115">Example</span></span>

<span data-ttu-id="fa25d-116">次の JavaScript のコード例は、Outlook の現在の項目の `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="fa25d-117">メンバー</span><span class="sxs-lookup"><span data-stu-id="fa25d-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="fa25d-118">添付ファイル :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa25d-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="fa25d-p103">項目の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-121">潜在的なセキュリティ問題により、特定の種類のファイルは Outlook でブロックされ、したがって戻ってきません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fa25d-122">詳細については、[「Outlook でブロックされた添付ファイル」](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="fa25d-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-123">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-123">Type:</span></span>

*   <span data-ttu-id="fa25d-124">配列.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fa25d-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-125">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-125">Requirements</span></span>

|<span data-ttu-id="fa25d-126">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-126">Requirement</span></span>| <span data-ttu-id="fa25d-127">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-128">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-128">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-129">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-129">1.0</span></span>|
|[<span data-ttu-id="fa25d-130">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-131">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-132">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-133">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-134">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-134">Example</span></span>

<span data-ttu-id="fa25d-135">次のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="fa25d-136">BCC:[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="fa25d-137">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fa25d-138">作成モード専用。</span><span class="sxs-lookup"><span data-stu-id="fa25d-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-139">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-139">Type:</span></span>

*   [<span data-ttu-id="fa25d-140">受信者</span><span class="sxs-lookup"><span data-stu-id="fa25d-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fa25d-141">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-141">Requirements</span></span>

|<span data-ttu-id="fa25d-142">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-142">Requirement</span></span>| <span data-ttu-id="fa25d-143">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-144">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-145">1.1以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-145">1.1</span></span>|
|[<span data-ttu-id="fa25d-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-147">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-149">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-150">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="fa25d-151">本文:[本文](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="fa25d-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="fa25d-152">項目の本文を操作するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-153">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-153">Type:</span></span>

*   [<span data-ttu-id="fa25d-154">本文</span><span class="sxs-lookup"><span data-stu-id="fa25d-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="fa25d-155">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-155">Requirements</span></span>

|<span data-ttu-id="fa25d-156">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-156">Requirement</span></span>| <span data-ttu-id="fa25d-157">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-158">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-159">1.1以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-159">1.1</span></span>|
|[<span data-ttu-id="fa25d-160">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-161">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-163">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="fa25d-164">CC:配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="fa25d-165">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fa25d-166">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-167">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-167">Read mode</span></span>

<span data-ttu-id="fa25d-p107">|||UNTRANSLATED_CONTENT_START|||The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="fa25d-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-170">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-170">Compose mode</span></span>

<span data-ttu-id="fa25d-171">|||UNTRANSLATED_CONTENT_START|||The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="fa25d-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-172">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-172">Type:</span></span>

*   <span data-ttu-id="fa25d-173">配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-174">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-174">Requirements</span></span>

|<span data-ttu-id="fa25d-175">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-175">Requirement</span></span>| <span data-ttu-id="fa25d-176">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-177">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-178">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-178">1.0</span></span>|
|[<span data-ttu-id="fa25d-179">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-180">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-181">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-182">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-183">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fa25d-184">（Null 許容）conversationId:文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="fa25d-185">特定のメッセージが含まれている電子メールの会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fa25d-p108">メール アプリを閲覧フォームでアクティブ化するか、新規作成フォームの返信でアクティブ化すると、このプロパティの整数を取得することができます。その後、ユーザーが返信の件名を変更した場合、その返信の送信時にメッセージの会話 ID が変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fa25d-p109">新規作成フォームで新しい項目に対してこのプロパティに null を取得します。ユーザーが件名を設定しアイテムを保存する場合、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-190">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-190">Type:</span></span>

*   <span data-ttu-id="fa25d-191">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-192">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-192">Requirements</span></span>

|<span data-ttu-id="fa25d-193">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-193">Requirement</span></span>| <span data-ttu-id="fa25d-194">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-195">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-195">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-196">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-196">1.0</span></span>|
|[<span data-ttu-id="fa25d-197">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-198">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-200">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="fa25d-201">dateTimeCreated:日付</span><span class="sxs-lookup"><span data-stu-id="fa25d-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="fa25d-p110">アイテムが作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-204">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-204">Type:</span></span>

*   <span data-ttu-id="fa25d-205">日付</span><span class="sxs-lookup"><span data-stu-id="fa25d-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-206">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-206">Requirements</span></span>

|<span data-ttu-id="fa25d-207">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-207">Requirement</span></span>| <span data-ttu-id="fa25d-208">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-209">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-209">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-210">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-210">1.0</span></span>|
|[<span data-ttu-id="fa25d-211">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-212">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-214">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-215">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fa25d-216">dateTimeModified:日付</span><span class="sxs-lookup"><span data-stu-id="fa25d-216">dateTimeModified :Date</span></span>

<span data-ttu-id="fa25d-p111">項目が最後に変更された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-219">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-220">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-220">Type:</span></span>

*   <span data-ttu-id="fa25d-221">日付</span><span class="sxs-lookup"><span data-stu-id="fa25d-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-222">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-222">Requirements</span></span>

|<span data-ttu-id="fa25d-223">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-223">Requirement</span></span>| <span data-ttu-id="fa25d-224">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-225">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-225">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-226">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-226">1.0</span></span>|
|[<span data-ttu-id="fa25d-227">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-228">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-230">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-231">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="fa25d-232">end:日付|[時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa25d-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="fa25d-233">予約が終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fa25d-p112">`end`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) メソッドを使用して、end プロパティ値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-236">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-236">Read mode</span></span>

<span data-ttu-id="fa25d-237">`end`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-238">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-238">Compose mode</span></span>

<span data-ttu-id="fa25d-239">`end`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fa25d-240">[`Time.setAsync` ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアントが所在するローカルの時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-241">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-241">Type:</span></span>

*   <span data-ttu-id="fa25d-242">日付 | [時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa25d-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-243">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-243">Requirements</span></span>

|<span data-ttu-id="fa25d-244">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-244">Requirement</span></span>| <span data-ttu-id="fa25d-245">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-246">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-246">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-247">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-247">1.0</span></span>|
|[<span data-ttu-id="fa25d-248">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-249">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-250">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-251">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-252">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-252">Example</span></span>

<span data-ttu-id="fa25d-253">次の例では、[ オブジェクトの`setAsync`  ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) `Time`  メソッドを使用して、作成モードで予定の終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="fa25d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fa25d-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="fa25d-p113">メッセージの送信者の電子メール アドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="fa25d-p114">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-259">`recipientType`  プロパティ内の`EmailAddressDetails`    オブジェクトの`from`   プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-260">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-260">Type:</span></span>

*   [<span data-ttu-id="fa25d-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fa25d-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fa25d-262">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-262">Requirements</span></span>

|<span data-ttu-id="fa25d-263">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-263">Requirement</span></span>| <span data-ttu-id="fa25d-264">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-265">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-265">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-266">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-266">1.0</span></span>|
|[<span data-ttu-id="fa25d-267">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-268">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-269">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-270">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="fa25d-271">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-271">internetMessageId :String</span></span>

<span data-ttu-id="fa25d-p115">電子メール メッセージ用のインターネット メッセージの識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-274">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-274">Type:</span></span>

*   <span data-ttu-id="fa25d-275">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-276">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-276">Requirements</span></span>

|<span data-ttu-id="fa25d-277">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-277">Requirement</span></span>| <span data-ttu-id="fa25d-278">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-279">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-279">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-280">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-280">1.0</span></span>|
|[<span data-ttu-id="fa25d-281">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-282">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-284">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-285">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fa25d-286">itemClass:文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-286">itemClass :String</span></span>

<span data-ttu-id="fa25d-p116">選択された項目の Exchange Web サービスの項目 クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fa25d-p117">`itemClass`プロパティは、選択したアイテムのメッセージ クラスを指定します。次は、メッセージまたは予定アイテムの既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="fa25d-291">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-291">Type</span></span> | <span data-ttu-id="fa25d-292">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-292">Description</span></span> | <span data-ttu-id="fa25d-293">項目 クラス</span><span class="sxs-lookup"><span data-stu-id="fa25d-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="fa25d-294">予定項目</span><span class="sxs-lookup"><span data-stu-id="fa25d-294">Appointment items</span></span> | <span data-ttu-id="fa25d-295">これらは、項目クラス `IPM.Appointment`または`IPM.Appointment.Occurence`の予定表アイテムです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="fa25d-296">メッセージの項目
</span><span class="sxs-lookup"><span data-stu-id="fa25d-296">Message items</span></span> | <span data-ttu-id="fa25d-297">これには、基本のメッセージ クラス として`IPM.Note`を使用する、既定のメッセージ クラス`IPM.Schedule.Meeting`会議出席依頼、返信および取り消しを含む電子メール メッセージが含まれます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="fa25d-298">既定のメッセージ クラスを拡張したカスタム メッセージ クラス 、たとえば、カスタム予定表メッセージ クラス`IPM.Appointment.Contoso` などを作成することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-299">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-299">Type:</span></span>

*   <span data-ttu-id="fa25d-300">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-301">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-301">Requirements</span></span>

|<span data-ttu-id="fa25d-302">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-302">Requirement</span></span>| <span data-ttu-id="fa25d-303">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-304">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-304">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-305">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-305">1.0</span></span>|
|[<span data-ttu-id="fa25d-306">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-307">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-308">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-309">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-310">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fa25d-311">(Null 許容)itemId: 文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-311">(nullable) itemId :String</span></span>

<span data-ttu-id="fa25d-p118">現在の項目の Exchange Web サービスのアイテム識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-314">`itemId` プロパティから返される識別子は、Exchange Web サービスの項目識別子と同じです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fa25d-315">`itemId` プロパティは、Outlook Entry ID または Outlook REST API によって使用される ID と同じではありません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fa25d-316">この値を使用して REST API の呼び出しを行う前に、 必要な設定1.3から利用可能な `Office.context.mailbox.convertToRestId`を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="fa25d-317">詳細については、[「Outlook アドインから Outlook REST API の使用」](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)をご覧下さい。</span><span class="sxs-lookup"><span data-stu-id="fa25d-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-318">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-318">Type:</span></span>

*   <span data-ttu-id="fa25d-319">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-320">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-320">Requirements</span></span>

|<span data-ttu-id="fa25d-321">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-321">Requirement</span></span>| <span data-ttu-id="fa25d-322">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-323">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-323">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-324">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-324">1.0</span></span>|
|[<span data-ttu-id="fa25d-325">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-326">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-327">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-328">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-329">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-329">Example</span></span>

<span data-ttu-id="fa25d-p120">次のコードは項目識別子のプレゼンスを確認します。`itemId` プロパティが `null` または `undefined` を返す場合、項目はストアに保存され、非同期の結果から項目識別子が取得されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="fa25d-332">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fa25d-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fa25d-333">インスタンスが表している項目の種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fa25d-334">`itemType`プロパティは、`ItemType`列挙値の 1 つを返します。これは`item`オブジェクト インスタンスがメッセージまたは予定のどちらであるかを示すものです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-335">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-335">Type:</span></span>

*   [<span data-ttu-id="fa25d-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fa25d-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fa25d-337">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-337">Requirements</span></span>

|<span data-ttu-id="fa25d-338">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-338">Requirement</span></span>| <span data-ttu-id="fa25d-339">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-340">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-340">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-341">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-341">1.0</span></span>|
|[<span data-ttu-id="fa25d-342">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-343">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-344">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-345">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-346">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="fa25d-347">場所: 文字列 | [場所](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="fa25d-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="fa25d-348">予約の場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-349">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-349">Read mode</span></span>

<span data-ttu-id="fa25d-350">`location`プロパティは、予定の場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-351">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-351">Compose mode</span></span>

<span data-ttu-id="fa25d-352">`location`プロパティは、予定の場所を取得または設定するために使用するメソッドを提供する`Location`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-353">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-353">Type:</span></span>

*   <span data-ttu-id="fa25d-354">文字列 | [場所](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="fa25d-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-355">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-355">Requirements</span></span>

|<span data-ttu-id="fa25d-356">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-356">Requirement</span></span>| <span data-ttu-id="fa25d-357">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-358">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-358">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-359">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-359">1.0</span></span>|
|[<span data-ttu-id="fa25d-360">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-361">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-362">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-363">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-364">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fa25d-365">normalizedSubject: 文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-365">normalizedSubject :String</span></span>

<span data-ttu-id="fa25d-p121">すべてのプレフィックス (`RE:` や `FWD:` を含む) が削除された項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fa25d-p122">normalizedSubject プロパティは、電子メール プログラムから追加された標準のプレフィックス (`RE:` や `FW:` など) が付く項目の件名を取得します。プレフィックスが付いたままの状態で項目の件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) プロパティを使用します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-370">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-370">Type:</span></span>

*   <span data-ttu-id="fa25d-371">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-372">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-372">Requirements</span></span>

|<span data-ttu-id="fa25d-373">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-373">Requirement</span></span>| <span data-ttu-id="fa25d-374">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-375">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-375">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-376">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-376">1.0</span></span>|
|[<span data-ttu-id="fa25d-377">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-378">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-379">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-380">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-381">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="fa25d-382">optionalAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="fa25d-383">オプションのイベント出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fa25d-384">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-385">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-385">Read mode</span></span>

<span data-ttu-id="fa25d-386">`optionalAttendees`プロパティは、会議への各任意出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-387">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-387">Compose mode</span></span>

<span data-ttu-id="fa25d-388">`optionalAttendees`プロパティは会議への任意出席者を取得または設定するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-389">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-389">Type:</span></span>

*   <span data-ttu-id="fa25d-390">配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-391">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-391">Requirements</span></span>

|<span data-ttu-id="fa25d-392">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-392">Requirement</span></span>| <span data-ttu-id="fa25d-393">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-394">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-394">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-395">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-395">1.0</span></span>|
|[<span data-ttu-id="fa25d-396">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-397">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-398">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-399">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-400">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="fa25d-401">開催者:[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fa25d-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="fa25d-p124">特定の会議開催者の電子メールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-404">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-404">Type:</span></span>

*   [<span data-ttu-id="fa25d-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fa25d-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fa25d-406">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-406">Requirements</span></span>

|<span data-ttu-id="fa25d-407">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-407">Requirement</span></span>| <span data-ttu-id="fa25d-408">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-409">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-409">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-410">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-410">1.0</span></span>|
|[<span data-ttu-id="fa25d-411">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-412">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-413">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-414">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-415">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="fa25d-416">requiredAttendees: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="fa25d-417">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fa25d-418">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-419">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-419">Read mode</span></span>

<span data-ttu-id="fa25d-420">`requiredAttendees`プロパティは、会議への各必須出席者の`EmailAddressDetails`オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-421">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-421">Compose mode</span></span>

<span data-ttu-id="fa25d-422">`requiredAttendees`プロパティは会議への必須出席者を取得または更新するためのメソッドを提供する`Recipients`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-423">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-423">Type:</span></span>

*   <span data-ttu-id="fa25d-424">配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-425">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-425">Requirements</span></span>

|<span data-ttu-id="fa25d-426">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-426">Requirement</span></span>| <span data-ttu-id="fa25d-427">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-428">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-428">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-429">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-429">1.0</span></span>|
|[<span data-ttu-id="fa25d-430">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-431">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-432">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-433">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-434">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="fa25d-435">送信者:[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fa25d-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="fa25d-p126">電子メール送信者のメールアドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fa25d-p127">メッセージが代理人から送信された場合を除き、[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails)プロパティと`sender`プロパティは同一人物を表します。その場合、`from`プロパティは委任者を、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-440">`recipientType`  プロパティ内の`EmailAddressDetails`    オブジェクトの`sender`   プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-441">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-441">Type:</span></span>

*   [<span data-ttu-id="fa25d-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fa25d-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fa25d-443">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-443">Requirements</span></span>

|<span data-ttu-id="fa25d-444">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-444">Requirement</span></span>| <span data-ttu-id="fa25d-445">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-446">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-446">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-447">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-447">1.0</span></span>|
|[<span data-ttu-id="fa25d-448">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-449">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-450">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-451">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-452">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="fa25d-453">開始: 日付 | [時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa25d-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="fa25d-454">予定を開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fa25d-p128">`start`プロパティは、協定世界時 (UTC) 形式の時刻値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime)メソッドを使用して、値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-457">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-457">Read mode</span></span>

<span data-ttu-id="fa25d-458">`start`プロパティは`Date`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-459">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-459">Compose mode</span></span>

<span data-ttu-id="fa25d-460">`start`プロパティは`Time`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fa25d-461">[`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)メソッドを使用して開始時刻を設定する場合、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアントが所在するローカル時刻をサーバーの UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-462">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-462">Type:</span></span>

*   <span data-ttu-id="fa25d-463">日付 | [時間](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="fa25d-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-464">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-464">Requirements</span></span>

|<span data-ttu-id="fa25d-465">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-465">Requirement</span></span>| <span data-ttu-id="fa25d-466">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-467">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-467">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-468">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-468">1.0</span></span>|
|[<span data-ttu-id="fa25d-469">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-470">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-471">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-472">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-473">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-473">Example</span></span>

<span data-ttu-id="fa25d-474">次の例では、[ オブジェクトの`setAsync` ](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-)  `Time`  メソッドを使用して、新規作成モードで予定の開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="fa25d-475">件名: 文字列 | [件名](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fa25d-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="fa25d-476">項目の件名フィールドに表示される説明を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fa25d-477">`subject`プロパティは、電子メールサーバーから送信された項目の全件名を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-478">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-478">Read mode</span></span>

<span data-ttu-id="fa25d-p129">`subject`プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string)プロパティを使用して、行間にある`RE:`や`FW:`のなどのすべてのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-481">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-481">Compose mode</span></span>

<span data-ttu-id="fa25d-482">`subject`プロパティは、件名を取得または設定するためのメソッドを提供する`Subject`オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fa25d-483">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-483">Type:</span></span>

*   <span data-ttu-id="fa25d-484">文字列 | [件名](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fa25d-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-485">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-485">Requirements</span></span>

|<span data-ttu-id="fa25d-486">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-486">Requirement</span></span>| <span data-ttu-id="fa25d-487">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-488">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-488">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-489">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-489">1.0</span></span>|
|[<span data-ttu-id="fa25d-490">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-491">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-492">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-493">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="fa25d-494">宛先: 配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="fa25d-495">メッセージの **宛先**列にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fa25d-496">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fa25d-497">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-497">Read mode</span></span>

<span data-ttu-id="fa25d-p131">`to`プロパティは、メッセージの `EmailAddressDetails` 宛先\*\*  列一覧にある各受信者の \*\*  オブジェクトを含む配列を返します。コレクションのメンバーは 100 個までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="fa25d-500">作成モード</span><span class="sxs-lookup"><span data-stu-id="fa25d-500">Compose mode</span></span>

<span data-ttu-id="fa25d-501">`to`プロパティは、メッセージの `Recipients` 宛先\*\*  列にある受信者を取得または更新するメソッドを提供する\*\* オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="fa25d-502">種類:</span><span class="sxs-lookup"><span data-stu-id="fa25d-502">Type:</span></span>

*   <span data-ttu-id="fa25d-503">配列.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>| [受信者](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fa25d-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-504">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-504">Requirements</span></span>

|<span data-ttu-id="fa25d-505">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-505">Requirement</span></span>| <span data-ttu-id="fa25d-506">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-507">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-507">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-508">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-508">1.0</span></span>|
|[<span data-ttu-id="fa25d-509">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-510">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-511">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-512">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-513">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="fa25d-514">メソッド</span><span class="sxs-lookup"><span data-stu-id="fa25d-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fa25d-515">addFileAttachmentAsync(uri, attachmentName, [オプション], [コールバック])</span><span class="sxs-lookup"><span data-stu-id="fa25d-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fa25d-516">ファイルを添付ファイルとしてメッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fa25d-517">`addFileAttachmentAsync`メソッドは、指定した URI にあるファイルをアップロードし、新規作成フォーム内の項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fa25d-518">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)メソッドで識別子を使用して同じセッションの添付ファイルを削除することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-519">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-519">Parameters:</span></span>

|<span data-ttu-id="fa25d-520">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-520">Name</span></span>| <span data-ttu-id="fa25d-521">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-521">Type</span></span>| <span data-ttu-id="fa25d-522">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-522">Attributes</span></span>| <span data-ttu-id="fa25d-523">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="fa25d-524">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-524">String</span></span>||<span data-ttu-id="fa25d-p132">メッセージまたは予定に添付するファイルの場所を示す URIです。 2048 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fa25d-527">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-527">String</span></span>||<span data-ttu-id="fa25d-p133">添付ファイルのアップロード時に表示される添付ファイルの名前です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fa25d-530">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-530">Object</span></span>| <span data-ttu-id="fa25d-531">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-531">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-532">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fa25d-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fa25d-533">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-533">Object</span></span>| <span data-ttu-id="fa25d-534">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-534">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-535">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fa25d-536">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-536">function</span></span>| <span data-ttu-id="fa25d-537">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-537">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-538">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa25d-539">これが成功すると、添付ファイルの識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fa25d-540">添付ファイルのアップロードに失敗した場合、`asyncResult`オブジェクトには、エラーの説明を提供する`Error`オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa25d-541">エラー</span><span class="sxs-lookup"><span data-stu-id="fa25d-541">Errors</span></span>

| <span data-ttu-id="fa25d-542">エラー コード</span><span class="sxs-lookup"><span data-stu-id="fa25d-542">Error code</span></span> | <span data-ttu-id="fa25d-543">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="fa25d-544">添付ファイルのサイズが大きすぎます。
</span><span class="sxs-lookup"><span data-stu-id="fa25d-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="fa25d-545">許可されていない拡張子の添付ファイルです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fa25d-546">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-547">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-547">Requirements</span></span>

|<span data-ttu-id="fa25d-548">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-548">Requirement</span></span>| <span data-ttu-id="fa25d-549">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-550">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-550">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-551">1.1以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-551">1.1</span></span>|
|[<span data-ttu-id="fa25d-552">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa25d-554">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-555">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-556">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-556">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fa25d-557">addItemAttachmentAsync (itemd、attachmentName、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="fa25d-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fa25d-558">メッセージなどの Exchange 項目を添付ファイルとして、メッセージまたは予定に追加します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fa25d-p134">`addItemAttachmentAsync` メソッドは、指定した Exchange 識別子を持つ項目を新規作成フォーム内の項目に添付します。コールバック メソッドを指定する場合、`asyncResult` という 1 つのパラメータがあるメソッドが呼び出されます。このパラメータには、添付ファイルの識別子、または項目の添付中に発生したエラーを示すコードが含まれます。必要に応じて、`options` パラメータを使用して、状態情報をコールバック メソッドに渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fa25d-562">その後、[`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback)メソッドで識別子を使用して同じセッションの添付ファイルを削除することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fa25d-563">Office アドインを Outlook Web アプリケーションで実行している場合、`addItemAttachmentAsync`メソッドは項目を、編集中以外の項目に添付できますが、これはサポートされておらず、推奨されていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-564">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-564">Parameters:</span></span>

|<span data-ttu-id="fa25d-565">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-565">Name</span></span>| <span data-ttu-id="fa25d-566">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-566">Type</span></span>| <span data-ttu-id="fa25d-567">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-567">Attributes</span></span>| <span data-ttu-id="fa25d-568">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="fa25d-569">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-569">String</span></span>||<span data-ttu-id="fa25d-p135">添付する項目の Exchange 識別子です。100 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fa25d-572">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-572">String</span></span>||<span data-ttu-id="fa25d-p136">添付する項目の件名です。 255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fa25d-575">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-575">Object</span></span>| <span data-ttu-id="fa25d-576">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-576">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-577">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fa25d-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fa25d-578">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-578">Object</span></span>| <span data-ttu-id="fa25d-579">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-579">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-580">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fa25d-581">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-581">function</span></span>| <span data-ttu-id="fa25d-582">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-582">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-583">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa25d-584">これが成功すると、添付ファイルの識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fa25d-585">添付ファイルの追加に失敗した場合、 エラーの説明を提供する`asyncResult`オブジェクトが`Error`オブジェクトに含まれます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa25d-586">エラー</span><span class="sxs-lookup"><span data-stu-id="fa25d-586">Errors</span></span>

| <span data-ttu-id="fa25d-587">エラー コード</span><span class="sxs-lookup"><span data-stu-id="fa25d-587">Error code</span></span> | <span data-ttu-id="fa25d-588">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fa25d-589">メッセージまたは予定の添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-590">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-590">Requirements</span></span>

|<span data-ttu-id="fa25d-591">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-591">Requirement</span></span>| <span data-ttu-id="fa25d-592">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-593">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-593">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-594">1.1以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-594">1.1</span></span>|
|[<span data-ttu-id="fa25d-595">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa25d-597">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-598">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-599">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-599">Example</span></span>

<span data-ttu-id="fa25d-600">次の例では、既存の Outlook 項目を名前付き `My Attachment` の添付ファイルとして追加されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="fa25d-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fa25d-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="fa25d-602">選択したメッセージの送信者とすべての受信者、または選択した予定の開催者とすべての参加者を示した返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-603">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa25d-604">Outlook Web アプリでは、返信フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fa25d-605">文字列パラメータのいずれかが制限値を超えると、`displayReplyAllForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="fa25d-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fa25d-p137">`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web アプリ はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-609">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-609">Parameters:</span></span>

|<span data-ttu-id="fa25d-610">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-610">Name</span></span>| <span data-ttu-id="fa25d-611">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-611">Type</span></span>| <span data-ttu-id="fa25d-612">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fa25d-613">文字列 | オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-613">String &#124; Object</span></span>| |<span data-ttu-id="fa25d-p138">返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fa25d-616">**または**</span><span class="sxs-lookup"><span data-stu-id="fa25d-616">**OR**</span></span><br/><span data-ttu-id="fa25d-p139">本文または添付ファイルのデータと、コールバック関数を含むオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fa25d-619">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-619">String</span></span> | <span data-ttu-id="fa25d-620">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-620">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-p140">返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fa25d-623">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fa25d-624">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-624">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-625">ファイルまたは項目の添付ファイルである JSON オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fa25d-626">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-626">String</span></span> | | <span data-ttu-id="fa25d-p141">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fa25d-629">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-629">String</span></span> | | <span data-ttu-id="fa25d-630">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fa25d-631">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-631">String</span></span> | | <span data-ttu-id="fa25d-p142">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fa25d-634">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-634">String</span></span> | | <span data-ttu-id="fa25d-p143">`type`が`item`に設定されている場合にのみ使用されます。添付ファイルの EWS アイテム ID です。 100 文字以内の文字列です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fa25d-638">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-638">function</span></span> | <span data-ttu-id="fa25d-639">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-639">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-640">メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-641">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-641">Requirements</span></span>

|<span data-ttu-id="fa25d-642">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-642">Requirement</span></span>| <span data-ttu-id="fa25d-643">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-644">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-645">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-645">1.0</span></span>|
|[<span data-ttu-id="fa25d-646">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-647">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-648">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-649">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa25d-650">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-650">Examples</span></span>

<span data-ttu-id="fa25d-651">次のコードは`displayReplyAllForm`関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fa25d-652">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fa25d-653">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fa25d-654">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-654">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fa25d-655">本文と項目の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-655">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fa25d-656">本文、添付ファイル、項目の添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="fa25d-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="fa25d-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="fa25d-658">選択したメッセージの送信者のみ、または選択した予定の開催者のみを含む返信フォームが表示されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-659">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa25d-660">Outlook Web アプリでは、返信フォームは、3 列ビューのポップアウト形式、および 2 列または 1 列ビューのポップアップ形式で表示されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fa25d-661">文字列パラメータのいずれかが制限値を超えると、`displayReplyForm` は例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="fa25d-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fa25d-p144">`formData.attachments` パラメータで添付ファイルを指定すると、Outlook と Office Web アプリ はすべての添付ファイルをダウンロードし、返信フォームに添付しようと試みます。添付ファイルの追加に失敗すると、フォーム UI でエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-665">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-665">Parameters:</span></span>

|<span data-ttu-id="fa25d-666">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-666">Name</span></span>| <span data-ttu-id="fa25d-667">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-667">Type</span></span>| <span data-ttu-id="fa25d-668">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="fa25d-669">文字列 | オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-669">String &#124; Object</span></span>| | <span data-ttu-id="fa25d-p145">返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fa25d-672">**または**</span><span class="sxs-lookup"><span data-stu-id="fa25d-672">**OR**</span></span><br/><span data-ttu-id="fa25d-p146">本文または添付ファイルのデータと、コールバック関数を含むオブジェクトです。オブジェクトの定義は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fa25d-675">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-675">String</span></span> | <span data-ttu-id="fa25d-676">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-676">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-p147">返信フォームの本文を表すテキストと HTML が含まれる文字列です。文字列は、32 KB 以内に制限されています。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fa25d-679">配列。&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fa25d-680">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-680">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-681">ファイルまたは項目の添付ファイルである JSON オブジェクトの配列です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fa25d-682">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-682">String</span></span> | | <span data-ttu-id="fa25d-p148">添付ファイルの種類を示します。添付ファイルの場合は`file`、添付項目の場合は`item`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fa25d-685">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-685">String</span></span> | | <span data-ttu-id="fa25d-686">添付ファイル名を含む文字列で、255 文字以内で入力が可能です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fa25d-687">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-687">String</span></span> | | <span data-ttu-id="fa25d-p149">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fa25d-690">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-690">String</span></span> | | <span data-ttu-id="fa25d-p150">`type`が`item`に設定されている場合にのみ使用されます。添付ファイルの EWS アイテム ID です。 100 文字以内の文字列です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fa25d-694">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-694">function</span></span> | <span data-ttu-id="fa25d-695">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-695">&lt;optional&gt;</span></span> | <span data-ttu-id="fa25d-696">メソッドが完了すると、 `callback` パラメータに渡された関数が、シングル パラメータ, `asyncResult`で呼び出されます。これは [AsyncResult](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-697">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-697">Requirements</span></span>

|<span data-ttu-id="fa25d-698">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-698">Requirement</span></span>| <span data-ttu-id="fa25d-699">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-700">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-700">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-701">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-701">1.0</span></span>|
|[<span data-ttu-id="fa25d-702">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-703">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-704">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-705">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fa25d-706">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-706">Examples</span></span>

<span data-ttu-id="fa25d-707">次のコードは`displayReplyForm`関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fa25d-708">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fa25d-709">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fa25d-710">本文と添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-710">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fa25d-711">本文と項目の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-711">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fa25d-712">本文、添付ファイル、項目の添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="fa25d-713">getEntities() → {[エンティティ](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fa25d-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="fa25d-714">選択した項目の本文で見つかったエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-715">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-716">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-716">Requirements</span></span>

|<span data-ttu-id="fa25d-717">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-717">Requirement</span></span>| <span data-ttu-id="fa25d-718">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-719">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-719">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-720">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-720">1.0</span></span>|
|[<span data-ttu-id="fa25d-721">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-722">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-723">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-724">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-725">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-725">Returns:</span></span>

<span data-ttu-id="fa25d-726">種類: [エンティティ](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fa25d-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fa25d-727">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-727">Example</span></span>

<span data-ttu-id="fa25d-728">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="fa25d-728">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="fa25d-729">getEntitiesByType(entityType)] → [(Null 許容) {配列<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="fa25d-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fa25d-730">選択した項目で見つかった指定のエンティティ型のエンティティすべてを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-731">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-732">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-732">Parameters:</span></span>

|<span data-ttu-id="fa25d-733">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-733">Name</span></span>| <span data-ttu-id="fa25d-734">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-734">Type</span></span>| <span data-ttu-id="fa25d-735">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="fa25d-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fa25d-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="fa25d-737">EntityType 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa25d-738">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-738">Requirements</span></span>

|<span data-ttu-id="fa25d-739">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-739">Requirement</span></span>| <span data-ttu-id="fa25d-740">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-741">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-741">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-742">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-742">1.0</span></span>|
|[<span data-ttu-id="fa25d-743">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-744">制限あり</span><span class="sxs-lookup"><span data-stu-id="fa25d-744">Restricted</span></span>|
|[<span data-ttu-id="fa25d-745">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-746">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-747">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-747">Returns:</span></span>

<span data-ttu-id="fa25d-748"> `entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは Nullを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fa25d-749">指定した種類のエンティティが項目の本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="fa25d-750">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fa25d-751">このメソッドを使用する最小限のアクセス許可のレベルは **制限あり**ですが、一部のエンティティには、次のテーブルで指定されているように、アクセスに **ReadItem** が必要です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="fa25d-752">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="fa25d-752">Value of `entityType`</span></span> | <span data-ttu-id="fa25d-753">返される配列内にあるオブジェクトの種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-753">Type of objects in returned array</span></span> | <span data-ttu-id="fa25d-754">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="fa25d-755">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-755">String</span></span> | <span data-ttu-id="fa25d-756">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="fa25d-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="fa25d-757">連絡先</span><span class="sxs-lookup"><span data-stu-id="fa25d-757">Contact</span></span> | <span data-ttu-id="fa25d-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa25d-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="fa25d-759">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-759">String</span></span> | <span data-ttu-id="fa25d-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa25d-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="fa25d-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fa25d-761">MeetingSuggestion</span></span> | <span data-ttu-id="fa25d-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa25d-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="fa25d-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fa25d-763">PhoneNumber</span></span> | <span data-ttu-id="fa25d-764">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="fa25d-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="fa25d-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fa25d-765">TaskSuggestion</span></span> | <span data-ttu-id="fa25d-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fa25d-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="fa25d-767">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-767">String</span></span> | <span data-ttu-id="fa25d-768">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="fa25d-768">**Restricted**</span></span> |

<span data-ttu-id="fa25d-769">種類: 配列.<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fa25d-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fa25d-770">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-770">Example</span></span>

<span data-ttu-id="fa25d-771">次の例は、現在の項目の本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="fa25d-772">getFilteredEntitiesByName(name)] → [(Null 許容) {配列<(文字列| [連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fa25d-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fa25d-773"> 選択済み項目の既知のエンティティを返し、 XML ファイルで定義された名前付きフィルターを渡します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-774">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa25d-775">`getFilteredEntitiesByName` メソッドは、指定された [    要素値があるマニフェストXMLファイル内のルール要素 ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule)  で定義された正規表現に一致するエンティティを返します。`FilterName`</span><span class="sxs-lookup"><span data-stu-id="fa25d-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-776">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-776">Parameters:</span></span>

|<span data-ttu-id="fa25d-777">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-777">Name</span></span>| <span data-ttu-id="fa25d-778">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-778">Type</span></span>| <span data-ttu-id="fa25d-779">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fa25d-780">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-780">String</span></span>|<span data-ttu-id="fa25d-781">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa25d-782">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-782">Requirements</span></span>

|<span data-ttu-id="fa25d-783">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-783">Requirement</span></span>| <span data-ttu-id="fa25d-784">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-785">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-785">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-786">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-786">1.0</span></span>|
|[<span data-ttu-id="fa25d-787">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-788">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-789">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-790">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-791">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-791">Returns:</span></span>

<span data-ttu-id="fa25d-p152"> `ItemHasKnownEntity`  パラメータと一致する `FilterName`  要素値を持つ `name`  要素がマニフェスト内にない場合、メソッドは `null\`を返します。  `name` パラメータがマニフェスト内の `ItemHasKnownEntity` 要素と一致するが、現在の項目内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fa25d-794">種類: 配列.<(文字列 |[連絡先](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fa25d-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="fa25d-795">getRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="fa25d-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fa25d-796">マニフェスト XML ファイルで定義された正規表現に一致する選択済みの項目の文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-797">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa25d-p153"> `getRegExMatches` メソッドは、マニフェスト XML ファイル内にある各 `ItemHasRegularExpressionMatch` または`ItemHasKnownEntity` ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致しなければなりません。 `PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fa25d-801">たとえば、アドイン マニフェストに次の `Rule` 要素があると見なします。</span><span class="sxs-lookup"><span data-stu-id="fa25d-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fa25d-802">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies`という 2 つのプロパティがあります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="fa25d-p154">項目の本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fa25d-805">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-805">Requirements</span></span>

|<span data-ttu-id="fa25d-806">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-806">Requirement</span></span>| <span data-ttu-id="fa25d-807">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-808">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-808">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-809">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-809">1.0</span></span>|
|[<span data-ttu-id="fa25d-810">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-811">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-812">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-813">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-814">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-814">Returns:</span></span>

<span data-ttu-id="fa25d-p155">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列を含むオブジェクトです。各配列の名前は、一致する `RegExName`  ルールの `ItemHasRegularExpressionMatch`  属性または一致する `FilterName`   ルールの `ItemHasKnownEntity`  属性の対応する値と等しくなります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fa25d-817">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa25d-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa25d-818">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa25d-819">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-819">Example</span></span>

<span data-ttu-id="fa25d-820">次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="fa25d-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fa25d-821">getRegExMatchesByName(name)] → [(Null許容) {配列. < 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="fa25d-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fa25d-822">選択した項目内の文字列を返し、マニフェスト XML ファイルで定義された名前付きの正規表現に一致します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fa25d-823">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fa25d-824">`getRegExMatchesByName` メソッドは、指定された`ItemHasRegularExpressionMatch`  要素値を持つマニフェスト XML ファイルの`RegExName`  ルール要素で定義された正規表現に一致する文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fa25d-p156">項目の本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-827">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-827">Parameters:</span></span>

|<span data-ttu-id="fa25d-828">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-828">Name</span></span>| <span data-ttu-id="fa25d-829">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-829">Type</span></span>| <span data-ttu-id="fa25d-830">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fa25d-831">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-831">String</span></span>|<span data-ttu-id="fa25d-832">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa25d-833">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-833">Requirements</span></span>

|<span data-ttu-id="fa25d-834">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-834">Requirement</span></span>| <span data-ttu-id="fa25d-835">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-836">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-836">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-837">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-837">1.0</span></span>|
|[<span data-ttu-id="fa25d-838">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-839">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-840">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-841">読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-842">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-842">Returns:</span></span>

<span data-ttu-id="fa25d-843">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="fa25d-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fa25d-844">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa25d-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa25d-845">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="fa25d-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa25d-846">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fa25d-847">getSelectedDataAsync(coercionType, [オプション], コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="fa25d-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fa25d-848">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fa25d-p157">選択したデータがなく、カーソルが本文または件名にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-851">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-851">Parameters:</span></span>

|<span data-ttu-id="fa25d-852">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-852">Name</span></span>| <span data-ttu-id="fa25d-853">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-853">Type</span></span>| <span data-ttu-id="fa25d-854">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-854">Attributes</span></span>| <span data-ttu-id="fa25d-855">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="fa25d-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fa25d-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fa25d-p158">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグはすべて削除されます。HTMLの場合、メソッドは、プレーン テキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="fa25d-860">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-860">Object</span></span>| <span data-ttu-id="fa25d-861">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-861">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-862">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fa25d-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fa25d-863">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-863">Object</span></span>| <span data-ttu-id="fa25d-864">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-864">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-865">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fa25d-866">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-866">function</span></span>||<span data-ttu-id="fa25d-867">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa25d-868">コールバック メソッドから選択したデータにアクセスするには、`asyncResult.value.data`を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fa25d-869">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`または `subject`になります。</span><span class="sxs-lookup"><span data-stu-id="fa25d-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa25d-870">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-870">Requirements</span></span>

|<span data-ttu-id="fa25d-871">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-871">Requirement</span></span>| <span data-ttu-id="fa25d-872">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-873">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-873">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-874">1.2以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-874">1.2</span></span>|
|[<span data-ttu-id="fa25d-875">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa25d-877">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-878">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fa25d-879">戻り値:</span><span class="sxs-lookup"><span data-stu-id="fa25d-879">Returns:</span></span>

<span data-ttu-id="fa25d-880">`coercionType`で決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fa25d-881">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="fa25d-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fa25d-882">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fa25d-883">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-883">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fa25d-884">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fa25d-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fa25d-885">選択された項目のアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fa25d-p160">カスタム プロパティは、アプリごと、項目ごとにキーと値のペアとして保管されます。このメソッドは、コールバックで  `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されないので、安全な保管場所として使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-889">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-889">Parameters:</span></span>

|<span data-ttu-id="fa25d-890">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-890">Name</span></span>| <span data-ttu-id="fa25d-891">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-891">Type</span></span>| <span data-ttu-id="fa25d-892">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-892">Attributes</span></span>| <span data-ttu-id="fa25d-893">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fa25d-894">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-894">function</span></span>||<span data-ttu-id="fa25d-895">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fa25d-896">カスタム プロパティは [  プロパティの `CustomProperties`  ](/javascript/api/outlook_1_2/office.customproperties) `asyncResult.value`  オブジェクトとして指定されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fa25d-897">このオブジェクトは、項目からカスタム プロパティを取得、設定、および削除し、カスタム プロパティに対する変更をサーバーに設定し直すために使用することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="fa25d-898">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-898">Object</span></span>| <span data-ttu-id="fa25d-899">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-899">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-900">開発者は、コールバック関数でアクセスしたいオブジェクトを提供することができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="fa25d-901">このオブジェクトは、コールバック関数の`asyncResult.asyncContext`プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fa25d-902">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-902">Requirements</span></span>

|<span data-ttu-id="fa25d-903">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-903">Requirement</span></span>| <span data-ttu-id="fa25d-904">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-905">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-905">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-906">1.0以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-906">1.0</span></span>|
|[<span data-ttu-id="fa25d-907">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-908">ReadItem</span></span>|
|[<span data-ttu-id="fa25d-909">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-910">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="fa25d-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-911">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-911">Example</span></span>

<span data-ttu-id="fa25d-p163">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、 `CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティを読み込んだ後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp`を読み取り、 `CustomProperties.set` メソッドでカスタム プロパティ `otherProp`を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fa25d-915">removeAttachmentAsync(attachmentId、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="fa25d-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fa25d-916">メッセージまたは予定から添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fa25d-p164">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-921">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-921">Parameters:</span></span>

|<span data-ttu-id="fa25d-922">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-922">Name</span></span>| <span data-ttu-id="fa25d-923">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-923">Type</span></span>| <span data-ttu-id="fa25d-924">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-924">Attributes</span></span>| <span data-ttu-id="fa25d-925">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="fa25d-926">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-926">String</span></span>||<span data-ttu-id="fa25d-p165">削除する添付ファイルの識別子です。配列は 100 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="fa25d-929">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-929">Object</span></span>| <span data-ttu-id="fa25d-930">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-930">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-931">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fa25d-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fa25d-932">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-932">Object</span></span>| <span data-ttu-id="fa25d-933">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-933">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-934">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fa25d-935">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-935">function</span></span>| <span data-ttu-id="fa25d-936">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-936">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-937">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fa25d-938">添付ファイルの削除に失敗すると、`asyncResult.error`プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fa25d-939">エラー</span><span class="sxs-lookup"><span data-stu-id="fa25d-939">Errors</span></span>

| <span data-ttu-id="fa25d-940">エラー コード</span><span class="sxs-lookup"><span data-stu-id="fa25d-940">Error code</span></span> | <span data-ttu-id="fa25d-941">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="fa25d-942">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="fa25d-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-943">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-943">Requirements</span></span>

|<span data-ttu-id="fa25d-944">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-944">Requirement</span></span>| <span data-ttu-id="fa25d-945">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-946">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-946">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-947">1.1以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-947">1.1</span></span>|
|[<span data-ttu-id="fa25d-948">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa25d-950">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-951">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-952">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-952">Example</span></span>

<span data-ttu-id="fa25d-953">次のコードは、「0」の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fa25d-954">setSelectedDataAsync(日付、 [オプション]、 コールバック)</span><span class="sxs-lookup"><span data-stu-id="fa25d-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fa25d-955">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="fa25d-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fa25d-p166">`setSelectedDataAsync`メソッドは、指定された文字列を項目のサブジェクトまたは本文のカーソル位置に挿入するか、エディターでテキストが選択されている場合は、選択されたテキストを置き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fa25d-959">パラメータ :</span><span class="sxs-lookup"><span data-stu-id="fa25d-959">Parameters:</span></span>

|<span data-ttu-id="fa25d-960">名前</span><span class="sxs-lookup"><span data-stu-id="fa25d-960">Name</span></span>| <span data-ttu-id="fa25d-961">種類</span><span class="sxs-lookup"><span data-stu-id="fa25d-961">Type</span></span>| <span data-ttu-id="fa25d-962">属性</span><span class="sxs-lookup"><span data-stu-id="fa25d-962">Attributes</span></span>| <span data-ttu-id="fa25d-963">説明</span><span class="sxs-lookup"><span data-stu-id="fa25d-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fa25d-964">文字列</span><span class="sxs-lookup"><span data-stu-id="fa25d-964">String</span></span>||<span data-ttu-id="fa25d-p167">挿入されるデータです。データの長さは 1,000,000 文字以内です。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="fa25d-968">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-968">Object</span></span>| <span data-ttu-id="fa25d-969">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-969">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-970">以下のプロパティのうち 1 つ以上を含むオブジェクト リテラル。</span><span class="sxs-lookup"><span data-stu-id="fa25d-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fa25d-971">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="fa25d-971">Object</span></span>| <span data-ttu-id="fa25d-972">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-972">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-973">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="fa25d-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fa25d-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="fa25d-975">&lt;オプション&gt;</span><span class="sxs-lookup"><span data-stu-id="fa25d-975">&lt;optional&gt;</span></span>|<span data-ttu-id="fa25d-p168"> `text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fa25d-p169"> `html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web アプリ では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、 `InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fa25d-980"> `coercionType` が設定されていない場合、結果はフィールドによって変わります。フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="fa25d-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="fa25d-981">関数</span><span class="sxs-lookup"><span data-stu-id="fa25d-981">function</span></span>||<span data-ttu-id="fa25d-982">メソッドが完了すると、`callback`パラメータで渡された関数が、シングル パラメータ、`asyncResult`、で呼び出されます。これは、[`AsyncResult`](/javascript/api/office/office.asyncresult)オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="fa25d-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fa25d-983">必要条件</span><span class="sxs-lookup"><span data-stu-id="fa25d-983">Requirements</span></span>

|<span data-ttu-id="fa25d-984">要件</span><span class="sxs-lookup"><span data-stu-id="fa25d-984">Requirement</span></span>| <span data-ttu-id="fa25d-985">値</span><span class="sxs-lookup"><span data-stu-id="fa25d-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="fa25d-986">メールボックスに必要な設定バージョン</span><span class="sxs-lookup"><span data-stu-id="fa25d-986">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fa25d-987">1.2以降</span><span class="sxs-lookup"><span data-stu-id="fa25d-987">1.2</span></span>|
|[<span data-ttu-id="fa25d-988">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fa25d-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fa25d-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fa25d-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="fa25d-990">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fa25d-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fa25d-991">作成</span><span class="sxs-lookup"><span data-stu-id="fa25d-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fa25d-992">例</span><span class="sxs-lookup"><span data-stu-id="fa25d-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```