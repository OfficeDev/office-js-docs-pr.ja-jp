
# <a name="item"></a><span data-ttu-id="85619-101">項目</span><span class="sxs-lookup"><span data-stu-id="85619-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="85619-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="85619-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="85619-p101">`item` ネームスペースを使用して、現在選択されているメッセージ、会議出席依頼、またはアポイントメントへアクセスします。[[項目タイプ（itemType）]](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) プロパティを使用して、`item` の種類を指定できます。</span><span class="sxs-lookup"><span data-stu-id="85619-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-105">要件</span><span class="sxs-lookup"><span data-stu-id="85619-105">Requirements</span></span>

|<span data-ttu-id="85619-106">要件</span><span class="sxs-lookup"><span data-stu-id="85619-106">Requirement</span></span>| <span data-ttu-id="85619-107">値</span><span class="sxs-lookup"><span data-stu-id="85619-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-108">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-109">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-109">1.0</span></span>|
|[<span data-ttu-id="85619-110">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-111">制限あり</span><span class="sxs-lookup"><span data-stu-id="85619-111">Restricted</span></span>|
|[<span data-ttu-id="85619-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="85619-114">例</span><span class="sxs-lookup"><span data-stu-id="85619-114">Example</span></span>

<span data-ttu-id="85619-115">以下の JavaScript コードの例は、Outlook の現在の項目の `subject` プロパティにアクセスする方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="85619-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="85619-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="85619-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="85619-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="85619-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="85619-p102">項目用の添付ファイルの配列を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-120">特定の種類のファイルは、潜在的なセキュリティ問題のため Outlook によりブロックされており、したがって返されません。</span><span class="sxs-lookup"><span data-stu-id="85619-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="85619-121">詳細については、「[Outlook でブロックされた添付ファイル](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="85619-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="85619-122">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-122">Type:</span></span>

*   <span data-ttu-id="85619-123">配列.<[[添付ファイルの詳細（AttachmentDetails）]](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="85619-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-124">要件</span><span class="sxs-lookup"><span data-stu-id="85619-124">Requirements</span></span>

|<span data-ttu-id="85619-125">要件</span><span class="sxs-lookup"><span data-stu-id="85619-125">Requirement</span></span>| <span data-ttu-id="85619-126">値</span><span class="sxs-lookup"><span data-stu-id="85619-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-127">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-127">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-128">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-128">1.0</span></span>|
|[<span data-ttu-id="85619-129">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-130">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-130">ReadItem</span></span>|
|[<span data-ttu-id="85619-131">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-132">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-133">例</span><span class="sxs-lookup"><span data-stu-id="85619-133">Example</span></span>

<span data-ttu-id="85619-134">以下のコードでは、現在の項目にあるすべての添付ファイルの詳細を含む HTML 文字列を作成します。</span><span class="sxs-lookup"><span data-stu-id="85619-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="85619-135">bcc:[受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="85619-136">メッセージの BCC (ブラインド カーボン コピー) 列上の 受信者を取得または更新するメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="85619-137">作成モードのみ。</span><span class="sxs-lookup"><span data-stu-id="85619-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-138">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-138">Type:</span></span>

*   [<span data-ttu-id="85619-139">受取者</span><span class="sxs-lookup"><span data-stu-id="85619-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="85619-140">要件</span><span class="sxs-lookup"><span data-stu-id="85619-140">Requirements</span></span>

|<span data-ttu-id="85619-141">要件</span><span class="sxs-lookup"><span data-stu-id="85619-141">Requirement</span></span>| <span data-ttu-id="85619-142">値</span><span class="sxs-lookup"><span data-stu-id="85619-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-143">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-143">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-144">1.1</span><span class="sxs-lookup"><span data-stu-id="85619-144">1.1</span></span>|
|[<span data-ttu-id="85619-145">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-146">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-146">ReadItem</span></span>|
|[<span data-ttu-id="85619-147">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-148">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-149">例</span><span class="sxs-lookup"><span data-stu-id="85619-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="85619-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="85619-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="85619-151">項目の本文を操作するためのメソッドを提供するオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-152">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-152">Type:</span></span>

*   [<span data-ttu-id="85619-153">本文</span><span class="sxs-lookup"><span data-stu-id="85619-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="85619-154">要件</span><span class="sxs-lookup"><span data-stu-id="85619-154">Requirements</span></span>

|<span data-ttu-id="85619-155">要件</span><span class="sxs-lookup"><span data-stu-id="85619-155">Requirement</span></span>| <span data-ttu-id="85619-156">値</span><span class="sxs-lookup"><span data-stu-id="85619-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-157">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-157">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-158">1.1</span><span class="sxs-lookup"><span data-stu-id="85619-158">1.1</span></span>|
|[<span data-ttu-id="85619-159">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-160">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-160">ReadItem</span></span>|
|[<span data-ttu-id="85619-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-162">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="85619-163">cc :配列.<[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="85619-164">メッセージの CC (カーボン コピー) 受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="85619-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="85619-165">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="85619-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-166">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-166">Read mode</span></span>

<span data-ttu-id="85619-p106">`cc` プロパティは、メッセージの **CC** 行上に一覧された各受信者の `EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは、最大 100 人のメンバーまでに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-169">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-169">Compose mode</span></span>

<span data-ttu-id="85619-170">`cc` プロパティは、メッセージの **CC** 行上の受信者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-171">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-171">Type:</span></span>

*   <span data-ttu-id="85619-172">配列.<[[E-メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-173">要件</span><span class="sxs-lookup"><span data-stu-id="85619-173">Requirements</span></span>

|<span data-ttu-id="85619-174">要件</span><span class="sxs-lookup"><span data-stu-id="85619-174">Requirement</span></span>| <span data-ttu-id="85619-175">値</span><span class="sxs-lookup"><span data-stu-id="85619-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-176">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-176">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-177">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-177">1.0</span></span>|
|[<span data-ttu-id="85619-178">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-179">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-179">ReadItem</span></span>|
|[<span data-ttu-id="85619-180">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-181">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-182">例</span><span class="sxs-lookup"><span data-stu-id="85619-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="85619-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="85619-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="85619-184">特定のメッセージを含む電子メール会話の識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="85619-p107">メール アプリが閲覧フォームでアクティブ化されている場合、または新規作成モードで応答する場合には、このプロパティに対して整数を取得することができます。その後にユーザーが返信メッセージの件名を変更した場合、返信を送信した時点で、そのメッセージの会話 ID は変更され、以前に取得した値は適用されなくなります。</span><span class="sxs-lookup"><span data-stu-id="85619-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="85619-p108">新規作成フォームの新しい項目については、このプロパティに対して Null を取得します。ユーザーが件名を設定して項目を保存すると、`conversationId` プロパティは値を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-189">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-189">Type:</span></span>

*   <span data-ttu-id="85619-190">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-191">要件</span><span class="sxs-lookup"><span data-stu-id="85619-191">Requirements</span></span>

|<span data-ttu-id="85619-192">要件</span><span class="sxs-lookup"><span data-stu-id="85619-192">Requirement</span></span>| <span data-ttu-id="85619-193">値</span><span class="sxs-lookup"><span data-stu-id="85619-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-194">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-195">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-195">1.0</span></span>|
|[<span data-ttu-id="85619-196">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-197">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-197">ReadItem</span></span>|
|[<span data-ttu-id="85619-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-199">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="85619-200">日時が作成されました：日付</span><span class="sxs-lookup"><span data-stu-id="85619-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="85619-p109">項目が作成された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-203">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-203">Type:</span></span>

*   <span data-ttu-id="85619-204">日付
</span><span class="sxs-lookup"><span data-stu-id="85619-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-205">要件</span><span class="sxs-lookup"><span data-stu-id="85619-205">Requirements</span></span>

|<span data-ttu-id="85619-206">要件</span><span class="sxs-lookup"><span data-stu-id="85619-206">Requirement</span></span>| <span data-ttu-id="85619-207">値</span><span class="sxs-lookup"><span data-stu-id="85619-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-208">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-208">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-209">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-209">1.0</span></span>|
|[<span data-ttu-id="85619-210">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-211">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-211">ReadItem</span></span>|
|[<span data-ttu-id="85619-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-213">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-214">例</span><span class="sxs-lookup"><span data-stu-id="85619-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="85619-215">日時が変更されました：日付</span><span class="sxs-lookup"><span data-stu-id="85619-215">dateTimeModified :Date</span></span>

<span data-ttu-id="85619-p110">項目が最後に変更された日時を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-218">このメンバーは、Outlook for iOS または Outlook for Android でサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-219">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-219">Type:</span></span>

*   <span data-ttu-id="85619-220">日付
</span><span class="sxs-lookup"><span data-stu-id="85619-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-221">要件</span><span class="sxs-lookup"><span data-stu-id="85619-221">Requirements</span></span>

|<span data-ttu-id="85619-222">要件</span><span class="sxs-lookup"><span data-stu-id="85619-222">Requirement</span></span>| <span data-ttu-id="85619-223">値</span><span class="sxs-lookup"><span data-stu-id="85619-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-224">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-225">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-225">1.0</span></span>|
|[<span data-ttu-id="85619-226">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-227">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-227">ReadItem</span></span>|
|[<span data-ttu-id="85619-228">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-229">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-230">例</span><span class="sxs-lookup"><span data-stu-id="85619-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="85619-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="85619-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="85619-232">アポイントメントが終了する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="85619-p111">`end` プロパティは、協定世界時 (UTC) の日時の値として表示されます。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) メソッドを使用して、 end プロパティの値をクライアントのローカル日時に変換することができます。</span><span class="sxs-lookup"><span data-stu-id="85619-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-235">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-235">Read mode</span></span>

<span data-ttu-id="85619-236">`end` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-237">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-237">Compose mode</span></span>

<span data-ttu-id="85619-238">`end` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="85619-239"> [ `Time.setAsync` ](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して終了時刻を設定する際には、 [ `convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカル時間をサーバー向けに UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="85619-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-240">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-240">Type:</span></span>

*   <span data-ttu-id="85619-241">日付 | [時間](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="85619-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-242">要件</span><span class="sxs-lookup"><span data-stu-id="85619-242">Requirements</span></span>

|<span data-ttu-id="85619-243">要件</span><span class="sxs-lookup"><span data-stu-id="85619-243">Requirement</span></span>| <span data-ttu-id="85619-244">値</span><span class="sxs-lookup"><span data-stu-id="85619-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-245">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-245">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-246">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-246">1.0</span></span>|
|[<span data-ttu-id="85619-247">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-248">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-248">ReadItem</span></span>|
|[<span data-ttu-id="85619-249">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-250">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-251">例</span><span class="sxs-lookup"><span data-stu-id="85619-251">Example</span></span>

<span data-ttu-id="85619-252">次の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードでアポイントメントの終了時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="85619-253">送信元：[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="85619-p112">メッセージの送信者の電子メール アドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="85619-p113">メッセージが代理人から送信された場合を除き、`from` プロパティと [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) プロパティは同一人物を表します。その場合、`from` プロパティは委任者を表し、送信者プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="85619-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-258">`from` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="85619-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-259">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-259">Type:</span></span>

*   <span data-ttu-id="85619-260">[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-260">[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-261">要件</span><span class="sxs-lookup"><span data-stu-id="85619-261">Requirements</span></span>

|<span data-ttu-id="85619-262">要件</span><span class="sxs-lookup"><span data-stu-id="85619-262">Requirement</span></span>| <span data-ttu-id="85619-263">値</span><span class="sxs-lookup"><span data-stu-id="85619-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-264">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-264">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-265">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-265">1.0</span></span>|
|[<span data-ttu-id="85619-266">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-267">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-267">ReadItem</span></span>|
|[<span data-ttu-id="85619-268">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-269">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="85619-270">internetMessageId: 文字列</span><span class="sxs-lookup"><span data-stu-id="85619-270">internetMessageId :String</span></span>

<span data-ttu-id="85619-p114">電子メール メッセージのインターネット メッセージ識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-273">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-273">Type:</span></span>

*   <span data-ttu-id="85619-274">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-275">要件</span><span class="sxs-lookup"><span data-stu-id="85619-275">Requirements</span></span>

|<span data-ttu-id="85619-276">要件</span><span class="sxs-lookup"><span data-stu-id="85619-276">Requirement</span></span>| <span data-ttu-id="85619-277">値</span><span class="sxs-lookup"><span data-stu-id="85619-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-278">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-279">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-279">1.0</span></span>|
|[<span data-ttu-id="85619-280">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-281">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-281">ReadItem</span></span>|
|[<span data-ttu-id="85619-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-283">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-284">例</span><span class="sxs-lookup"><span data-stu-id="85619-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="85619-285">itemClass: 文字列</span><span class="sxs-lookup"><span data-stu-id="85619-285">itemClass :String</span></span>

<span data-ttu-id="85619-p115">選択された項目の Exchange Web サービスの項目クラスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="85619-p116">`itemClass` プロパティは、選択された項目のメッセージ クラスを指定します。以下は、メッセージまたはアポイントメント項目の既定のメッセージ クラスを示しています。</span><span class="sxs-lookup"><span data-stu-id="85619-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="85619-290">種類</span><span class="sxs-lookup"><span data-stu-id="85619-290">Type</span></span> | <span data-ttu-id="85619-291">記述</span><span class="sxs-lookup"><span data-stu-id="85619-291">Description</span></span> | <span data-ttu-id="85619-292">項目クラス</span><span class="sxs-lookup"><span data-stu-id="85619-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="85619-293">アポイントメント項目</span><span class="sxs-lookup"><span data-stu-id="85619-293">Appointment items</span></span> | <span data-ttu-id="85619-294">これらは、項目クラス `IPM.Appointment` または `IPM.Appointment.Occurence` のカレンダー項目です。</span><span class="sxs-lookup"><span data-stu-id="85619-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="85619-295">メッセージ項目</span><span class="sxs-lookup"><span data-stu-id="85619-295">Message items</span></span> | <span data-ttu-id="85619-296">これらには、既定のメッセージ クラス `IPM.Note` を持つ電子メール メッセージと、`IPM.Schedule.Meeting` を基本のメッセージ クラスとして使用する会議出席依頼、応答、キャンセルが含まれます。</span><span class="sxs-lookup"><span data-stu-id="85619-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="85619-297">既定のメッセージ クラスを拡張するカスタム メッセージ クラス (たとえば、カスタムのアポイントメント メッセージ クラス `IPM.Appointment.Contoso` など) を作成することができます。</span><span class="sxs-lookup"><span data-stu-id="85619-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-298">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-298">Type:</span></span>

*   <span data-ttu-id="85619-299">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-300">要件</span><span class="sxs-lookup"><span data-stu-id="85619-300">Requirements</span></span>

|<span data-ttu-id="85619-301">要件</span><span class="sxs-lookup"><span data-stu-id="85619-301">Requirement</span></span>| <span data-ttu-id="85619-302">値</span><span class="sxs-lookup"><span data-stu-id="85619-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-303">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-303">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-304">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-304">1.0</span></span>|
|[<span data-ttu-id="85619-305">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-306">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-306">ReadItem</span></span>|
|[<span data-ttu-id="85619-307">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-308">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-309">例</span><span class="sxs-lookup"><span data-stu-id="85619-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="85619-310">(Null可能) [項目ID（itemId）]：文字列</span><span class="sxs-lookup"><span data-stu-id="85619-310">(nullable) itemId :String</span></span>

<span data-ttu-id="85619-p117">現在の項目の Exchange Web サービスの項目識別子を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-313">`itemId` プロパティによって返される識別子は、Exchange Web サービスの項目識別子と同じものです。</span><span class="sxs-lookup"><span data-stu-id="85619-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="85619-314">`itemId` プロパティは、Outlook Entry ID または Outlook REST API によって使用される ID と同じものではありません。</span><span class="sxs-lookup"><span data-stu-id="85619-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="85619-315">この値を使用して REST の API 呼び出しを行う前に、[Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) を使用して変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="85619-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="85619-316">詳細については、「[Outlook アドインから Outlook REST API を使用する](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="85619-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="85619-p119">`itemId` プロパティは、新規作成モードで利用できません。項目識別子が必要な場合には、[`saveAsync`](#saveasyncoptions-callback) メソッドを使用してストアに項目を保存することができます。こうすることで、コールバック関数の [`AsyncResult.value`](/javascript/api/office/office.asyncresult) パラメータで項目識別子が返されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-319">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-319">Type:</span></span>

*   <span data-ttu-id="85619-320">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-321">要件</span><span class="sxs-lookup"><span data-stu-id="85619-321">Requirements</span></span>

|<span data-ttu-id="85619-322">要件</span><span class="sxs-lookup"><span data-stu-id="85619-322">Requirement</span></span>| <span data-ttu-id="85619-323">値</span><span class="sxs-lookup"><span data-stu-id="85619-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-324">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-324">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-325">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-325">1.0</span></span>|
|[<span data-ttu-id="85619-326">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-327">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-327">ReadItem</span></span>|
|[<span data-ttu-id="85619-328">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-329">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-330">例</span><span class="sxs-lookup"><span data-stu-id="85619-330">Example</span></span>

<span data-ttu-id="85619-p120">以下のコードは、項目識別子の有無を確認します。`itemId` プロパティが `null` または `undefined` を返す場合には、このコードは項目をストアに保存し、非同期の結果から項目識別子を取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="85619-333">[項目の種類（itemType）]：[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="85619-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="85619-334">インスタンスが表す項目の種類を取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="85619-335">`itemType` プロパティは、 `ItemType` 列挙の内の 1 つ値を返します。この値は、`item` オブジェクト インスタンスがメッセージであるか、それともアポイントメントであるかを示します。</span><span class="sxs-lookup"><span data-stu-id="85619-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-336">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-336">Type:</span></span>

*   [<span data-ttu-id="85619-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="85619-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="85619-338">要件</span><span class="sxs-lookup"><span data-stu-id="85619-338">Requirements</span></span>

|<span data-ttu-id="85619-339">要件</span><span class="sxs-lookup"><span data-stu-id="85619-339">Requirement</span></span>| <span data-ttu-id="85619-340">値</span><span class="sxs-lookup"><span data-stu-id="85619-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-341">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-341">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-342">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-342">1.0</span></span>|
|[<span data-ttu-id="85619-343">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-344">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-344">ReadItem</span></span>|
|[<span data-ttu-id="85619-345">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-346">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-347">例</span><span class="sxs-lookup"><span data-stu-id="85619-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="85619-348">場所：文字列|[場所](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="85619-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="85619-349">アポイントメントの場所を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-350">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-350">Read mode</span></span>

<span data-ttu-id="85619-351">`location` プロパティは、アポイントメントの場所を含む文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-352">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-352">Compose mode</span></span>

<span data-ttu-id="85619-353">`location` プロパティは、アポイントメントの場所を取得または設定するために利用されるメソッドを提供する `Location` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-354">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-354">Type:</span></span>

*   <span data-ttu-id="85619-355">文字列 | [場所](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="85619-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-356">要件</span><span class="sxs-lookup"><span data-stu-id="85619-356">Requirements</span></span>

|<span data-ttu-id="85619-357">要件</span><span class="sxs-lookup"><span data-stu-id="85619-357">Requirement</span></span>| <span data-ttu-id="85619-358">値</span><span class="sxs-lookup"><span data-stu-id="85619-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-359">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-359">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-360">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-360">1.0</span></span>|
|[<span data-ttu-id="85619-361">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-362">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-362">ReadItem</span></span>|
|[<span data-ttu-id="85619-363">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-364">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-365">例</span><span class="sxs-lookup"><span data-stu-id="85619-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="85619-366">normalizedSubject :文字列</span><span class="sxs-lookup"><span data-stu-id="85619-366">normalizedSubject :String</span></span>

<span data-ttu-id="85619-p121">すべてのプレフィックス (`RE:`および`FWD:` を含む) を削除した、項目の件名を取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="85619-p122">normalizedSubject プロパティは、電子メール プログラムにより追加された標準のプレフィックス (`RE:`および`FW:` など) とともに、項目の件名を取得します。プレフィックスが付いたままの状態で項目の件名を取得するには、[`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) プロパティを使用してください。</span><span class="sxs-lookup"><span data-stu-id="85619-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-371">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-371">Type:</span></span>

*   <span data-ttu-id="85619-372">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-373">要件</span><span class="sxs-lookup"><span data-stu-id="85619-373">Requirements</span></span>

|<span data-ttu-id="85619-374">要件</span><span class="sxs-lookup"><span data-stu-id="85619-374">Requirement</span></span>| <span data-ttu-id="85619-375">値</span><span class="sxs-lookup"><span data-stu-id="85619-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-376">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-376">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-377">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-377">1.0</span></span>|
|[<span data-ttu-id="85619-378">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-379">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-379">ReadItem</span></span>|
|[<span data-ttu-id="85619-380">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-381">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-382">例</span><span class="sxs-lookup"><span data-stu-id="85619-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="85619-383">[通知メッセージ（notificationMessages）]：[[通知メッセージ（notificationMessages）]](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="85619-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="85619-384">項目の通知メッセージを取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-385">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-385">Type:</span></span>

*   [<span data-ttu-id="85619-386">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="85619-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="85619-387">要件</span><span class="sxs-lookup"><span data-stu-id="85619-387">Requirements</span></span>

|<span data-ttu-id="85619-388">要件</span><span class="sxs-lookup"><span data-stu-id="85619-388">Requirement</span></span>| <span data-ttu-id="85619-389">値</span><span class="sxs-lookup"><span data-stu-id="85619-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-390">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-390">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-391">1.3</span><span class="sxs-lookup"><span data-stu-id="85619-391">1.3</span></span>|
|[<span data-ttu-id="85619-392">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-393">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-393">ReadItem</span></span>|
|[<span data-ttu-id="85619-394">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-395">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="85619-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="85619-397">イベントへの任意の出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="85619-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="85619-398">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="85619-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-399">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-399">Read mode</span></span>

<span data-ttu-id="85619-400">`optionalAttendees` プロパティは、会議の各任意出席者に対して `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-401">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-401">Compose mode</span></span>

<span data-ttu-id="85619-402">`optionalAttendees` プロパティは会議の任意出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-403">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-403">Type:</span></span>

*   <span data-ttu-id="85619-404">配列.<[[E-メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-405">要件</span><span class="sxs-lookup"><span data-stu-id="85619-405">Requirements</span></span>

|<span data-ttu-id="85619-406">要件</span><span class="sxs-lookup"><span data-stu-id="85619-406">Requirement</span></span>| <span data-ttu-id="85619-407">値</span><span class="sxs-lookup"><span data-stu-id="85619-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-408">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-408">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-409">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-409">1.0</span></span>|
|[<span data-ttu-id="85619-410">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-411">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-411">ReadItem</span></span>|
|[<span data-ttu-id="85619-412">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-413">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-414">例</span><span class="sxs-lookup"><span data-stu-id="85619-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="85619-415">開催者:[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="85619-p124">指定の会議の開催者の電子メール アドレスを取得します。閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-418">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-418">Type:</span></span>

*   <span data-ttu-id="85619-419">[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-419">[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-420">要件</span><span class="sxs-lookup"><span data-stu-id="85619-420">Requirements</span></span>

|<span data-ttu-id="85619-421">要件</span><span class="sxs-lookup"><span data-stu-id="85619-421">Requirement</span></span>| <span data-ttu-id="85619-422">値</span><span class="sxs-lookup"><span data-stu-id="85619-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-423">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-423">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-424">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-424">1.0</span></span>|
|[<span data-ttu-id="85619-425">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-426">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-426">ReadItem</span></span>|
|[<span data-ttu-id="85619-427">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-428">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-429">例</span><span class="sxs-lookup"><span data-stu-id="85619-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="85619-430">[必須出席者（requiredAttendees）]：配列 。<[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_4/office.recipients)。</span><span class="sxs-lookup"><span data-stu-id="85619-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="85619-431">イベントの必須出席者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="85619-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="85619-432">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="85619-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-433">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-433">Read mode</span></span>

<span data-ttu-id="85619-434">`requiredAttendees` プロパティは、会議の各必須出席者に対して `EmailAddressDetails` オブジェクトを含む配列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-435">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-435">Compose mode</span></span>

<span data-ttu-id="85619-436">`requiredAttendees` プロパティは、会議への必須出席者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-437">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-437">Type:</span></span>

*   <span data-ttu-id="85619-438">配列.<[[E-メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-439">要件</span><span class="sxs-lookup"><span data-stu-id="85619-439">Requirements</span></span>

|<span data-ttu-id="85619-440">要件</span><span class="sxs-lookup"><span data-stu-id="85619-440">Requirement</span></span>| <span data-ttu-id="85619-441">値</span><span class="sxs-lookup"><span data-stu-id="85619-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-442">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-443">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-443">1.0</span></span>|
|[<span data-ttu-id="85619-444">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-445">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-445">ReadItem</span></span>|
|[<span data-ttu-id="85619-446">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-447">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-448">例</span><span class="sxs-lookup"><span data-stu-id="85619-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="85619-449">送信者：[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="85619-p126">電子メール メッセージの送信者の電子メール アドレスを取得します、閲覧モード専用です。</span><span class="sxs-lookup"><span data-stu-id="85619-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="85619-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) および `sender` プロパティは、代理人によりメッセージが送信された場合を除き、同一人物を表します。その場合、`from` プロパティは委任者を表し、送信者 プロパティは代理人を表します。</span><span class="sxs-lookup"><span data-stu-id="85619-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-454">`sender` プロパティ内の`EmailAddressDetails` オブジェクトの`recipientType`  プロパティは、 `undefined`です。</span><span class="sxs-lookup"><span data-stu-id="85619-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-455">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-455">Type:</span></span>

*   <span data-ttu-id="85619-456">[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="85619-456">[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-457">要件</span><span class="sxs-lookup"><span data-stu-id="85619-457">Requirements</span></span>

|<span data-ttu-id="85619-458">要件</span><span class="sxs-lookup"><span data-stu-id="85619-458">Requirement</span></span>| <span data-ttu-id="85619-459">値</span><span class="sxs-lookup"><span data-stu-id="85619-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-460">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-460">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-461">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-461">1.0</span></span>|
|[<span data-ttu-id="85619-462">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-463">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-463">ReadItem</span></span>|
|[<span data-ttu-id="85619-464">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-465">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-466">例</span><span class="sxs-lookup"><span data-stu-id="85619-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="85619-467">開始：日付|[時間](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="85619-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="85619-468">アポイントメントが開始する日時を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="85619-p128">`start` プロパティは、協定世界時 (UTC) の値で日時を表します。[`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) メソッドを使用して、この値をクライアントのローカル日時に変換する事ができます。</span><span class="sxs-lookup"><span data-stu-id="85619-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-471">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-471">Read mode</span></span>

<span data-ttu-id="85619-472">`start` プロパティは `Date` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-473">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-473">Compose mode</span></span>

<span data-ttu-id="85619-474">`start` プロパティは `Time` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="85619-475">[ `Time.setAsync` ](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して開始時刻を設定する際には、[`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) メソッドを使用して、クライアント上のローカル時刻をサーバー向けに UTC に変換する必要があります。</span><span class="sxs-lookup"><span data-stu-id="85619-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-476">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-476">Type:</span></span>

*   <span data-ttu-id="85619-477">日付 | [時間](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="85619-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-478">要件</span><span class="sxs-lookup"><span data-stu-id="85619-478">Requirements</span></span>

|<span data-ttu-id="85619-479">要件</span><span class="sxs-lookup"><span data-stu-id="85619-479">Requirement</span></span>| <span data-ttu-id="85619-480">値</span><span class="sxs-lookup"><span data-stu-id="85619-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-481">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-481">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-482">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-482">1.0</span></span>|
|[<span data-ttu-id="85619-483">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-484">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-484">ReadItem</span></span>|
|[<span data-ttu-id="85619-485">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-486">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-487">例</span><span class="sxs-lookup"><span data-stu-id="85619-487">Example</span></span>

<span data-ttu-id="85619-488">以下の例では、`Time` オブジェクトの [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) メソッドを使用して、新規作成モードでアポイントメントの開始時刻を設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="85619-489">件名：文字列|[件名](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="85619-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="85619-490">項目の件名フィールドに表示される記述を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="85619-491">`subject` プロパティは、電子メール サーバーから送信された通りに、項目の件名全体を取得または設定します。</span><span class="sxs-lookup"><span data-stu-id="85619-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-492">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-492">Read mode</span></span>

<span data-ttu-id="85619-p129">`subject` プロパティは文字列を返します。[`normalizedSubject`](#normalizedsubject-string) プロパティを使用して、`RE:`や`FW:` など行間にある全てのプレフィックスを削除した件名を取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="85619-495">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-495">Compose mode</span></span>

<span data-ttu-id="85619-496">`subject` プロパティは、件名を取得または設定するためのメソッドを提供する `Subject` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="85619-497">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-497">Type:</span></span>

*   <span data-ttu-id="85619-498">文字列 | [件名](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="85619-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-499">要件</span><span class="sxs-lookup"><span data-stu-id="85619-499">Requirements</span></span>

|<span data-ttu-id="85619-500">要件</span><span class="sxs-lookup"><span data-stu-id="85619-500">Requirement</span></span>| <span data-ttu-id="85619-501">値</span><span class="sxs-lookup"><span data-stu-id="85619-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-502">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-502">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-503">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-503">1.0</span></span>|
|[<span data-ttu-id="85619-504">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-505">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-505">ReadItem</span></span>|
|[<span data-ttu-id="85619-506">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-507">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="85619-508">宛先：配列.<[[電子メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="85619-509">メッセージの **宛先** 行にある受信者へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="85619-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="85619-510">オブジェクトの種類とアクセスのレベルは、現在の項目のモードによって異なります。</span><span class="sxs-lookup"><span data-stu-id="85619-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="85619-511">閲覧モード</span><span class="sxs-lookup"><span data-stu-id="85619-511">Read mode</span></span>

<span data-ttu-id="85619-p131">`to` プロパティは、メッセージの**宛先**行に一覧された各受信者の `EmailAddressDetails` オブジェクトを含む配列を返します。コレクションは、最大 100 人のメンバーに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="85619-514">新規作成モード</span><span class="sxs-lookup"><span data-stu-id="85619-514">Compose mode</span></span>

<span data-ttu-id="85619-515">`to` プロパティは、メッセージの**宛先**行にある受信者を取得または更新するためのメソッドを提供する `Recipients` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="85619-516">種類:</span><span class="sxs-lookup"><span data-stu-id="85619-516">Type:</span></span>

*   <span data-ttu-id="85619-517">配列.<[[E-メールアドレスの詳細（EmailAddressDetails）]](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [受信者](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="85619-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-518">要件</span><span class="sxs-lookup"><span data-stu-id="85619-518">Requirements</span></span>

|<span data-ttu-id="85619-519">要件</span><span class="sxs-lookup"><span data-stu-id="85619-519">Requirement</span></span>| <span data-ttu-id="85619-520">値</span><span class="sxs-lookup"><span data-stu-id="85619-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-521">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-522">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-522">1.0</span></span>|
|[<span data-ttu-id="85619-523">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-524">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-524">ReadItem</span></span>|
|[<span data-ttu-id="85619-525">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-526">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-527">例</span><span class="sxs-lookup"><span data-stu-id="85619-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="85619-528">メソッド</span><span class="sxs-lookup"><span data-stu-id="85619-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="85619-529">[非同期の添付ファイルを追加（addFileAttachmentAsync）](uri, [添付ファイル名（attachmentName）], [オプション（options）], [コールバック（callback）])</span><span class="sxs-lookup"><span data-stu-id="85619-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="85619-530">ファイルを添付ファイルとして、メッセージまたはアポイントメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="85619-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="85619-531">`addFileAttachmentAsync` メソッドは、指定した URL にあるファイルをアップロードし、そのファイルを新規作成フォームの項目に添付します。</span><span class="sxs-lookup"><span data-stu-id="85619-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="85619-532">その後に、 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して、同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="85619-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-533">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-533">Parameters:</span></span>

|<span data-ttu-id="85619-534">名前</span><span class="sxs-lookup"><span data-stu-id="85619-534">Name</span></span>| <span data-ttu-id="85619-535">種類</span><span class="sxs-lookup"><span data-stu-id="85619-535">Type</span></span>| <span data-ttu-id="85619-536">特性</span><span class="sxs-lookup"><span data-stu-id="85619-536">Attributes</span></span>| <span data-ttu-id="85619-537">記述</span><span class="sxs-lookup"><span data-stu-id="85619-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="85619-538">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-538">String</span></span>||<span data-ttu-id="85619-p132">メッセージまたはアポイントメントに添付するファイルの場所を提供する URL です。2048 文字が最大の文字数です。</span><span class="sxs-lookup"><span data-stu-id="85619-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="85619-541">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-541">String</span></span>||<span data-ttu-id="85619-p133">添付ファイルのアップロード中に表示される、その添付ファイルの名前です。255 文字が最大の文字数です。</span><span class="sxs-lookup"><span data-stu-id="85619-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="85619-544">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-544">Object</span></span>| <span data-ttu-id="85619-545">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-545">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-546">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-547">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-547">Object</span></span>| <span data-ttu-id="85619-548">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-548">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-549">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="85619-550">関数</span><span class="sxs-lookup"><span data-stu-id="85619-550">function</span></span>| <span data-ttu-id="85619-551">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-551">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-552">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="85619-553">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティにて提供されます。</span><span class="sxs-lookup"><span data-stu-id="85619-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="85619-554">添付ファイルのアップロードが失敗した場合、`asyncResult` オブジェクトには、エラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="85619-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="85619-555">エラー</span><span class="sxs-lookup"><span data-stu-id="85619-555">Errors</span></span>

| <span data-ttu-id="85619-556">エラー コード</span><span class="sxs-lookup"><span data-stu-id="85619-556">Error code</span></span> | <span data-ttu-id="85619-557">記述</span><span class="sxs-lookup"><span data-stu-id="85619-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="85619-558">添付ファイルのサイズが上限を超えています。</span><span class="sxs-lookup"><span data-stu-id="85619-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="85619-559">添付ファイルに許可されていない拡張機能が含まれています。</span><span class="sxs-lookup"><span data-stu-id="85619-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="85619-560">メッセージまたはアポイントメントの添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="85619-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-561">要件</span><span class="sxs-lookup"><span data-stu-id="85619-561">Requirements</span></span>

|<span data-ttu-id="85619-562">要件</span><span class="sxs-lookup"><span data-stu-id="85619-562">Requirement</span></span>| <span data-ttu-id="85619-563">値</span><span class="sxs-lookup"><span data-stu-id="85619-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-564">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-565">1.1</span><span class="sxs-lookup"><span data-stu-id="85619-565">1.1</span></span>|
|[<span data-ttu-id="85619-566">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-567">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-568">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-569">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-570">例</span><span class="sxs-lookup"><span data-stu-id="85619-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="85619-571">[非同期の項目の添付ファイルを追加（addItemAttachmentAsync）] ([項目ID（itemId）]、[添付ファイル名（attachmentName）]、[オプション]、[コールバック])</span><span class="sxs-lookup"><span data-stu-id="85619-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="85619-572">メッセージなど Exchange の項目を、メッセージまたはアポイントメントの添付ファイルとして追加します。</span><span class="sxs-lookup"><span data-stu-id="85619-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="85619-p134">`addItemAttachmentAsync` メソッドは、指定された Exchange 識別子を持つ項目を、新規作成フォームの項目に添付します。コールバック メソッドを指定した場合、このメソッドは 1 つのパラメータ `asyncResult` で呼び出されます。このパラメータには、添付ファイルの識別子または項目を添付する間に発生したエラーを表示するコードが含まれます。必要に応じて、`options` パラメータを使用してコールバック メソッドに状態に関する情報を渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="85619-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="85619-576">その後に、 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) メソッドで識別子を使用して、同じセッションの添付ファイルを削除できます。</span><span class="sxs-lookup"><span data-stu-id="85619-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="85619-577">Outlook Web App で Office アドインを実行している場合、 `addItemAttachmentAsync` メソッドで、編集しているもの以外の項目に項目を添付する事ができます。ただし、この操作はサポートされておらず、お勧めしません。</span><span class="sxs-lookup"><span data-stu-id="85619-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-578">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-578">Parameters:</span></span>

|<span data-ttu-id="85619-579">名前</span><span class="sxs-lookup"><span data-stu-id="85619-579">Name</span></span>| <span data-ttu-id="85619-580">種類</span><span class="sxs-lookup"><span data-stu-id="85619-580">Type</span></span>| <span data-ttu-id="85619-581">特性</span><span class="sxs-lookup"><span data-stu-id="85619-581">Attributes</span></span>| <span data-ttu-id="85619-582">記述</span><span class="sxs-lookup"><span data-stu-id="85619-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="85619-583">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-583">String</span></span>||<span data-ttu-id="85619-p135">添付する項目の Exchange 識別子。100 文字以内で入力してください</span><span class="sxs-lookup"><span data-stu-id="85619-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="85619-586">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-586">String</span></span>||<span data-ttu-id="85619-p136">添付する項目の件名。255 文字以内で入力してください。</span><span class="sxs-lookup"><span data-stu-id="85619-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="85619-589">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-589">Object</span></span>| <span data-ttu-id="85619-590">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-590">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-591">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-592">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-592">Object</span></span>| <span data-ttu-id="85619-593">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-593">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-594">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="85619-595">関数</span><span class="sxs-lookup"><span data-stu-id="85619-595">function</span></span>| <span data-ttu-id="85619-596">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-596">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-597">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="85619-598">成功すると、添付ファイルの識別子が `asyncResult.value` プロパティにて提供されます。</span><span class="sxs-lookup"><span data-stu-id="85619-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="85619-599">ファイルの添付が失敗すると、`asyncResult` オブジェクトにはエラーの説明を提供する `Error` オブジェクトが含まれます。</span><span class="sxs-lookup"><span data-stu-id="85619-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="85619-600">エラー</span><span class="sxs-lookup"><span data-stu-id="85619-600">Errors</span></span>

| <span data-ttu-id="85619-601">エラー コード</span><span class="sxs-lookup"><span data-stu-id="85619-601">Error code</span></span> | <span data-ttu-id="85619-602">記述</span><span class="sxs-lookup"><span data-stu-id="85619-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="85619-603">メッセージまたはアポイントメントの添付ファイルが多すぎます。</span><span class="sxs-lookup"><span data-stu-id="85619-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-604">要件</span><span class="sxs-lookup"><span data-stu-id="85619-604">Requirements</span></span>

|<span data-ttu-id="85619-605">要件</span><span class="sxs-lookup"><span data-stu-id="85619-605">Requirement</span></span>| <span data-ttu-id="85619-606">値</span><span class="sxs-lookup"><span data-stu-id="85619-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-607">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-607">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-608">1.1</span><span class="sxs-lookup"><span data-stu-id="85619-608">1.1</span></span>|
|[<span data-ttu-id="85619-609">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-610">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-611">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-612">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-613">例</span><span class="sxs-lookup"><span data-stu-id="85619-613">Example</span></span>

<span data-ttu-id="85619-614">以下の例では、既存の Outlook の項目を、`My Attachment` という名前の添付ファイルとして追加します。</span><span class="sxs-lookup"><span data-stu-id="85619-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="85619-615">閉じる()</span><span class="sxs-lookup"><span data-stu-id="85619-615">close()</span></span>

<span data-ttu-id="85619-616">新規作成中の現在の項目を閉じます。</span><span class="sxs-lookup"><span data-stu-id="85619-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="85619-p137">`close` メソッドの動作は新規作成中の項目の現在の状態によって異なります。項目に保存されていない変更があると、クライアントはユーザーに「閉じる」操作を保存、破棄、またはキャンセルするように指示します。</span><span class="sxs-lookup"><span data-stu-id="85619-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-619">Outlook on the web では、項目がアポイントメントであり、以前に `saveAsync` を使用して保存されていた場合、項目が最後に保存されてから何も変更されていなくても、ユーザーは保存、破棄、またはキャンセルをするように指示されます。</span><span class="sxs-lookup"><span data-stu-id="85619-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="85619-620">Outlook デスクトップ クライアント内では、メッセージがインラインの返信の場合、`close` メソッドには何の効果もありません。</span><span class="sxs-lookup"><span data-stu-id="85619-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-621">要件</span><span class="sxs-lookup"><span data-stu-id="85619-621">Requirements</span></span>

|<span data-ttu-id="85619-622">要件</span><span class="sxs-lookup"><span data-stu-id="85619-622">Requirement</span></span>| <span data-ttu-id="85619-623">値</span><span class="sxs-lookup"><span data-stu-id="85619-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-624">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-624">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-625">1.3</span><span class="sxs-lookup"><span data-stu-id="85619-625">1.3</span></span>|
|[<span data-ttu-id="85619-626">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-627">制限あり</span><span class="sxs-lookup"><span data-stu-id="85619-627">Restricted</span></span>|
|[<span data-ttu-id="85619-628">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-629">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="85619-630">[全返信フォームを表示（フォームデータ）（displayReplyAllForm(formData)）]</span><span class="sxs-lookup"><span data-stu-id="85619-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="85619-631">選択されたメッセージの送信者とすべての受信者、または選択されたアポイントメントの主催者とすべての出席者を含む返信フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="85619-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-632">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="85619-633">Outlook Web App では、返信フォームは、3 列表示のポップアップアウト フォームと、2 列または 1 列表示のポップアップ フォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="85619-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="85619-634">文字列パラメータのいずれかが制限値を超えると、`displayReplyAllForm` が例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="85619-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="85619-p138">添付ファイルが `formData.attachments` パラメータで指定されている場合、Outlook と Outlook Web App は、すべての添付ファイルをダウンロードして、それらを返信フォームに添付しようとします。いずれかの添付ファイルの追加に失敗すると、フォーム IU にエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="85619-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-638">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-638">Parameters:</span></span>

|<span data-ttu-id="85619-639">名前</span><span class="sxs-lookup"><span data-stu-id="85619-639">Name</span></span>| <span data-ttu-id="85619-640">種類</span><span class="sxs-lookup"><span data-stu-id="85619-640">Type</span></span>| <span data-ttu-id="85619-641">記述</span><span class="sxs-lookup"><span data-stu-id="85619-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="85619-642">文字列 | オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-642">String &#124; Object</span></span>| |<span data-ttu-id="85619-p139">テキストと HTML を含み、返信フォームの本文を表す文字列。この文字列は 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="85619-645">**または**</span><span class="sxs-lookup"><span data-stu-id="85619-645">**OR**</span></span><br/><span data-ttu-id="85619-p140">本文または添付ファイルのデータと、コールバック関数を含むオブジェクト。このオブジェクトは次のように定義されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="85619-648">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-648">String</span></span> | <span data-ttu-id="85619-649">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-649">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-p141">テキストと HTML を含み、返信フォームの本文を表す文字列。この文字列は 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="85619-652">配列.&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="85619-653">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-653">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-654">ファイルまたは項目のいずれかの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="85619-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="85619-655">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-655">String</span></span> | | <span data-ttu-id="85619-p142">添付ファイルの種類を示します。ファイルの添付の場合は `file`で、項目の添付の場合は `item` でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="85619-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="85619-658">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-658">String</span></span> | | <span data-ttu-id="85619-659">添付ファイルの名前を含む文字列。最大文字数は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="85619-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="85619-660">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-660">String</span></span> | | <span data-ttu-id="85619-p143">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="85619-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="85619-663">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-663">String</span></span> | | <span data-ttu-id="85619-p144">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="85619-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="85619-667">関数</span><span class="sxs-lookup"><span data-stu-id="85619-667">function</span></span> | <span data-ttu-id="85619-668">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-668">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-669">メソッドが完了すると、`callback` パラメーターに渡された関数は、[[非同期の結果（AsyncResult）]](/javascript/api/office/office.asyncresult) オブジェクトである 単一 パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="85619-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-670">要件</span><span class="sxs-lookup"><span data-stu-id="85619-670">Requirements</span></span>

|<span data-ttu-id="85619-671">要件</span><span class="sxs-lookup"><span data-stu-id="85619-671">Requirement</span></span>| <span data-ttu-id="85619-672">値</span><span class="sxs-lookup"><span data-stu-id="85619-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-673">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-673">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-674">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-674">1.0</span></span>|
|[<span data-ttu-id="85619-675">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-676">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-676">ReadItem</span></span>|
|[<span data-ttu-id="85619-677">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-678">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="85619-679">例</span><span class="sxs-lookup"><span data-stu-id="85619-679">Examples</span></span>

<span data-ttu-id="85619-680">次のコードは `displayReplyAllForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="85619-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="85619-681">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="85619-682">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="85619-683">本文と「ファイル」の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="85619-684">本文と「項目」の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="85619-685">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="85619-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="85619-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="85619-687">選択したメッセージの送信者のみ、または選択したアポイントメントの開催者のみを含む返信フォームを表示します。</span><span class="sxs-lookup"><span data-stu-id="85619-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-688">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="85619-689">Outlook Web App では、返信フォームは、3 列表示のポップアップアウト フォームと、2 列または 1 列表示のポップアップ フォームとして表示されます。</span><span class="sxs-lookup"><span data-stu-id="85619-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="85619-690">文字列パラメータのいずれかが制限値を超えると、`displayReplyForm` が例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="85619-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="85619-p145">添付ファイルが `formData.attachments` パラメータで指定されている場合、Outlook と Outlook Web App は、すべての添付ファイルをダウンロードして、それらを返信フォームに添付しようとします。いずれかの添付ファイルの追加に失敗すると、フォーム IU にエラーが表示されます。表示できない場合、エラー メッセージはスローされません。</span><span class="sxs-lookup"><span data-stu-id="85619-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-694">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-694">Parameters:</span></span>

|<span data-ttu-id="85619-695">名前</span><span class="sxs-lookup"><span data-stu-id="85619-695">Name</span></span>| <span data-ttu-id="85619-696">種類</span><span class="sxs-lookup"><span data-stu-id="85619-696">Type</span></span>| <span data-ttu-id="85619-697">記述</span><span class="sxs-lookup"><span data-stu-id="85619-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="85619-698">文字列 | オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-698">String &#124; Object</span></span>| | <span data-ttu-id="85619-p146">テキストと HTML を含み、返信フォームの本文を表す文字列。この文字列は 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="85619-701">**または**</span><span class="sxs-lookup"><span data-stu-id="85619-701">**OR**</span></span><br/><span data-ttu-id="85619-p147">本文または添付ファイルのデータと、コールバック関数を含むオブジェクト。このオブジェクトは次のように定義されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="85619-704">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-704">String</span></span> | <span data-ttu-id="85619-705">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-705">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-p148">テキストと HTML を含み、返信フォームの本文を表す文字列。この文字列は 32 KB までに制限されています。</span><span class="sxs-lookup"><span data-stu-id="85619-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="85619-708">配列.&lt;オブジェクト&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="85619-709">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-709">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-710">ファイルまたは項目のいずれかの添付ファイルである JSON オブジェクトの配列。</span><span class="sxs-lookup"><span data-stu-id="85619-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="85619-711">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-711">String</span></span> | | <span data-ttu-id="85619-p149">添付ファイルの種類を示します。ファイルの添付の場合は `file`で、項目の添付の場合は `item` でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="85619-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="85619-714">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-714">String</span></span> | | <span data-ttu-id="85619-715">添付ファイルの名前を含む文字列。最大文字数は 255 文字です。</span><span class="sxs-lookup"><span data-stu-id="85619-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="85619-716">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-716">String</span></span> | | <span data-ttu-id="85619-p150">`type` が `file` に設定されている場合にのみ使用されます。ファイルの場所の URIです。</span><span class="sxs-lookup"><span data-stu-id="85619-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="85619-719">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-719">String</span></span> | | <span data-ttu-id="85619-p151">`type` が `item` に設定されている場合にのみ使用されます。添付ファイルの EWS アイテムIDです。最大 100 文字の文字列です。</span><span class="sxs-lookup"><span data-stu-id="85619-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="85619-723">関数</span><span class="sxs-lookup"><span data-stu-id="85619-723">function</span></span> | <span data-ttu-id="85619-724">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-724">&lt;optional&gt;</span></span> | <span data-ttu-id="85619-725">メソッドが完了すると、`callback` パラメーターに渡された関数は、[[非同期の結果（AsyncResult）]](/javascript/api/office/office.asyncresult) オブジェクトである 単一 パラメーター `asyncResult` で呼び出されます。</span><span class="sxs-lookup"><span data-stu-id="85619-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-726">要件</span><span class="sxs-lookup"><span data-stu-id="85619-726">Requirements</span></span>

|<span data-ttu-id="85619-727">要件</span><span class="sxs-lookup"><span data-stu-id="85619-727">Requirement</span></span>| <span data-ttu-id="85619-728">値</span><span class="sxs-lookup"><span data-stu-id="85619-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-729">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-729">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-730">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-730">1.0</span></span>|
|[<span data-ttu-id="85619-731">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-732">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-732">ReadItem</span></span>|
|[<span data-ttu-id="85619-733">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-734">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="85619-735">例</span><span class="sxs-lookup"><span data-stu-id="85619-735">Examples</span></span>

<span data-ttu-id="85619-736">次のコードは `displayReplyForm` 関数に文字列を渡します。</span><span class="sxs-lookup"><span data-stu-id="85619-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="85619-737">空の本文を返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="85619-738">本文だけを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="85619-739">本文と「ファイル」の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="85619-740">本文と「項目」の添付ファイルを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="85619-741">本文、ファイルの添付ファイル、アイテムの添付ファイル、およびコールバックを返信します。</span><span class="sxs-lookup"><span data-stu-id="85619-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="85619-742">[エンティティを取得（getEntities）]() → {[エンティティ](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="85619-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="85619-743">選択した項目の本文にあるエンティティを取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-744">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-745">要件</span><span class="sxs-lookup"><span data-stu-id="85619-745">Requirements</span></span>

|<span data-ttu-id="85619-746">要件</span><span class="sxs-lookup"><span data-stu-id="85619-746">Requirement</span></span>| <span data-ttu-id="85619-747">値</span><span class="sxs-lookup"><span data-stu-id="85619-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-748">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-748">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-749">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-749">1.0</span></span>|
|[<span data-ttu-id="85619-750">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-751">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-751">ReadItem</span></span>|
|[<span data-ttu-id="85619-752">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-753">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-754">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-754">Returns:</span></span>

<span data-ttu-id="85619-755">種類: [エンティティ](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="85619-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="85619-756">例</span><span class="sxs-lookup"><span data-stu-id="85619-756">Example</span></span>

<span data-ttu-id="85619-757">次の例では、現在の項目の本文内にある連絡先のエンティティにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="85619-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="85619-758">[種類別でエンティティを取得（エンティティの種類）（getEntitiesByType(entityType)）] → [(Null可能) {<(文字列|[連絡先](/javascript/api/outlook_1_4/office.contact)|[[会議の提案（MeetingSuggestion）]](/javascript/api/outlook_1_4/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_4/office.phonenumber)|[[タスクの提案（TaskSuggestion）]](/javascript/api/outlook_1_4/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="85619-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="85619-759">選択した項目の本文で見つかった指定のエンティティの種類の全てのエンティティを含む配列を取得します。</span><span class="sxs-lookup"><span data-stu-id="85619-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-760">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-761">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-761">Parameters:</span></span>

|<span data-ttu-id="85619-762">名前</span><span class="sxs-lookup"><span data-stu-id="85619-762">Name</span></span>| <span data-ttu-id="85619-763">種類</span><span class="sxs-lookup"><span data-stu-id="85619-763">Type</span></span>| <span data-ttu-id="85619-764">記述</span><span class="sxs-lookup"><span data-stu-id="85619-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="85619-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="85619-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="85619-766">[エンティティの種類（EntityType）] 列挙値の 1 つです。</span><span class="sxs-lookup"><span data-stu-id="85619-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-767">要件</span><span class="sxs-lookup"><span data-stu-id="85619-767">Requirements</span></span>

|<span data-ttu-id="85619-768">要件</span><span class="sxs-lookup"><span data-stu-id="85619-768">Requirement</span></span>| <span data-ttu-id="85619-769">値</span><span class="sxs-lookup"><span data-stu-id="85619-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-770">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-770">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-771">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-771">1.0</span></span>|
|[<span data-ttu-id="85619-772">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-773">制限あり</span><span class="sxs-lookup"><span data-stu-id="85619-773">Restricted</span></span>|
|[<span data-ttu-id="85619-774">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-775">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-776">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-776">Returns:</span></span>

<span data-ttu-id="85619-777">`entityType` に渡された値が有効な `EntityType` 列挙型のメンバーでない場合、メソッドは 空白を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="85619-778">指定した種類のエンティティが項目の本文に存在しない場合、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="85619-779">そうでない場合、返される配列内のオブジェクトの種類は、 `entityType` パラメータ内で要求されたエンティティの種類によって異なります。</span><span class="sxs-lookup"><span data-stu-id="85619-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="85619-780">このメソッドを使用する最小限のアクセス許可レベルは **制限あり** ですが、一部のエンティティ型には、次の表で指定されているように、アクセスに **[項目の閲覧（ReadItem）]** が必要です。</span><span class="sxs-lookup"><span data-stu-id="85619-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="85619-781">の値 `entityType`</span><span class="sxs-lookup"><span data-stu-id="85619-781">Value of `entityType`</span></span> | <span data-ttu-id="85619-782">返される配列内のオブジェクトの種類</span><span class="sxs-lookup"><span data-stu-id="85619-782">Type of objects in returned array</span></span> | <span data-ttu-id="85619-783">必要なアクセス許可のレベル</span><span class="sxs-lookup"><span data-stu-id="85619-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="85619-784">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-784">String</span></span> | <span data-ttu-id="85619-785">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="85619-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="85619-786">連絡先</span><span class="sxs-lookup"><span data-stu-id="85619-786">Contact</span></span> | <span data-ttu-id="85619-787">**項目を閲覧**</span><span class="sxs-lookup"><span data-stu-id="85619-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="85619-788">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-788">String</span></span> | <span data-ttu-id="85619-789">**項目を閲覧**</span><span class="sxs-lookup"><span data-stu-id="85619-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="85619-790">[会議の提案（MeetingSuggestion）]</span><span class="sxs-lookup"><span data-stu-id="85619-790">MeetingSuggestion</span></span> | <span data-ttu-id="85619-791">**項目を閲覧**</span><span class="sxs-lookup"><span data-stu-id="85619-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="85619-792">電話番号</span><span class="sxs-lookup"><span data-stu-id="85619-792">PhoneNumber</span></span> | <span data-ttu-id="85619-793">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="85619-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="85619-794">[タスクの提案（TaskSuggestion）]</span><span class="sxs-lookup"><span data-stu-id="85619-794">TaskSuggestion</span></span> | <span data-ttu-id="85619-795">**項目を閲覧**</span><span class="sxs-lookup"><span data-stu-id="85619-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="85619-796">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-796">String</span></span> | <span data-ttu-id="85619-797">**制限あり**</span><span class="sxs-lookup"><span data-stu-id="85619-797">**Restricted**</span></span> |

<span data-ttu-id="85619-798">種類：配列.<(文字列|[連絡先](/javascript/api/outlook_1_4/office.contact)|[[会議の提案（MeetingSuggestion）]](/javascript/api/outlook_1_4/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_4/office.phonenumber)|[[タスクの提案（TaskSuggestion）]](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="85619-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="85619-799">例</span><span class="sxs-lookup"><span data-stu-id="85619-799">Example</span></span>

<span data-ttu-id="85619-800">次の例は、現在の項目の本文にある郵便アドレスを表す文字列の配列にアクセスする方法を示します。</span><span class="sxs-lookup"><span data-stu-id="85619-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="85619-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(文字列| [連絡先](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[電話番号 ](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="85619-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="85619-802">マニフェスト XML ファイルで定義された名前付きフィルターを通過する、選択したアイテム内の既知のエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-803">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="85619-804">`getFilteredEntitiesByName` メソッドは、マニフェスト XML ファイル内の、指定された `FilterName` 要素値を持つ [[項目は知られたエンティティを持つ（ItemHasKnownEntity）]](/javascript/office/manifest/rule#itemhasknownentity-rule) ルール要素で定義された正規表現に一致するエンティティを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-805">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-805">Parameters:</span></span>

|<span data-ttu-id="85619-806">名前</span><span class="sxs-lookup"><span data-stu-id="85619-806">Name</span></span>| <span data-ttu-id="85619-807">種類</span><span class="sxs-lookup"><span data-stu-id="85619-807">Type</span></span>| <span data-ttu-id="85619-808">記述</span><span class="sxs-lookup"><span data-stu-id="85619-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="85619-809">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-809">String</span></span>|<span data-ttu-id="85619-810">一致するフィルターを定義する `ItemHasKnownEntity` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="85619-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-811">要件</span><span class="sxs-lookup"><span data-stu-id="85619-811">Requirements</span></span>

|<span data-ttu-id="85619-812">要件</span><span class="sxs-lookup"><span data-stu-id="85619-812">Requirement</span></span>| <span data-ttu-id="85619-813">値</span><span class="sxs-lookup"><span data-stu-id="85619-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-814">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-814">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-815">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-815">1.0</span></span>|
|[<span data-ttu-id="85619-816">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-817">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-817">ReadItem</span></span>|
|[<span data-ttu-id="85619-818">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-819">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-820">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-820">Returns:</span></span>

<span data-ttu-id="85619-p153">マニフェスト内に`FilterName`  要素の値が `name` パラメーターと一致する `ItemHasKnownEntity` 要素がない場合、メソッドは `null` を返します。`name` パラメーターがマニフェスト内の `ItemHasKnownEntity` 要素と一致しながら、現在のアイテム内に一致するエンティティがない場合は、メソッドは空の配列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="85619-823">種類：配列.<(文字列|[連絡先](/javascript/api/outlook_1_4/office.contact)|[[会議の提案（MeetingSuggestion）]](/javascript/api/outlook_1_4/office.meetingsuggestion)|[電話番号](/javascript/api/outlook_1_4/office.phonenumber)|[[タスクの提案（TaskSuggestion）]](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="85619-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="85619-824">getRegExMatches() → {オブジェクト}</span><span class="sxs-lookup"><span data-stu-id="85619-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="85619-825">選択した項目内で、マニフェスト XML ファイルで定義された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-826">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="85619-p154">`getRegExMatches` メソッドは、マニフェスト XML ファイル内にある `ItemHasRegularExpressionMatch` または `ItemHasKnownEntity` の各ルール要素で定義された正規表現に一致する文字列を返します。`ItemHasRegularExpressionMatch` ルールの場合、そのルールで指定された項目のプロパティに一致する文字列が発生する必要があります。`PropertyName` 単純型は、サポートされるプロパティを定義します。</span><span class="sxs-lookup"><span data-stu-id="85619-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="85619-830">たとえば、アドイン マニフェストに次のような `Rule` 要素があると見なします：</span><span class="sxs-lookup"><span data-stu-id="85619-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="85619-831">`getRegExMatches` から返されるオブジェクトには、 `fruits` および `veggies` という 2 つのプロパティを持つ可能性があります。</span><span class="sxs-lookup"><span data-stu-id="85619-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="85619-p155">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用してアイテムの本文全体を取得しても、期待する結果が返されないことがあります。代わりに [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) メソッドを使用して本文全体を取得してください。</span><span class="sxs-lookup"><span data-stu-id="85619-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="85619-835">要件</span><span class="sxs-lookup"><span data-stu-id="85619-835">Requirements</span></span>

|<span data-ttu-id="85619-836">要件</span><span class="sxs-lookup"><span data-stu-id="85619-836">Requirement</span></span>| <span data-ttu-id="85619-837">値</span><span class="sxs-lookup"><span data-stu-id="85619-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-838">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-838">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-839">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-839">1.0</span></span>|
|[<span data-ttu-id="85619-840">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-841">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-841">ReadItem</span></span>|
|[<span data-ttu-id="85619-842">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-843">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-844">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-844">Returns:</span></span>

<span data-ttu-id="85619-p156">マニフェスト XML ファイルで定義された正規表現に一致する文字列の配列が格納されたオブジェクトです。各配列の名前は、一致する `ItemHasRegularExpressionMatch` ルールの `RegExName` 属性、または一致する `ItemHasKnownEntity` ルールの `FilterName` 属性に対応する値に等しいものとなります。</span><span class="sxs-lookup"><span data-stu-id="85619-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="85619-847">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="85619-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="85619-848">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="85619-849">例</span><span class="sxs-lookup"><span data-stu-id="85619-849">Example</span></span>

<span data-ttu-id="85619-850">次の例は、マニフェストで指定された正規表現の <rule> 要素 `fruits` および `veggies` に一致する配列にアクセスする方法を示しています。</rule></span><span class="sxs-lookup"><span data-stu-id="85619-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="85619-851">getRegExMatchesByName(name)] → (nullable) {Array.< 文字列 >}</span><span class="sxs-lookup"><span data-stu-id="85619-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="85619-852">選択した項目内の、マニフェスト XML ファイルで定義された、指定された正規表現に一致する文字列の値を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-853">このメソッドは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="85619-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="85619-854">`getRegExMatchesByName` メソッドは、指定された `RegExName` 要素の値を持つマニフェスト XML ファイル内にある`ItemHasRegularExpressionMatch` ルール要素で定義された正規表現に一致する文字列を返します。</span><span class="sxs-lookup"><span data-stu-id="85619-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="85619-p157">項目の 本文プロパティに `ItemHasRegularExpressionMatch` ルールを指定する場合、正規表現でさらに本文をフィルター処理し、項目の本文全体を返さないようにします。`.*` などの正規表現を使用して項目の本文全体を取得しても、期待する結果が返されないことがあります。</span><span class="sxs-lookup"><span data-stu-id="85619-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-857">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-857">Parameters:</span></span>

|<span data-ttu-id="85619-858">名前</span><span class="sxs-lookup"><span data-stu-id="85619-858">Name</span></span>| <span data-ttu-id="85619-859">種類</span><span class="sxs-lookup"><span data-stu-id="85619-859">Type</span></span>| <span data-ttu-id="85619-860">記述</span><span class="sxs-lookup"><span data-stu-id="85619-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="85619-861">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-861">String</span></span>|<span data-ttu-id="85619-862">一致するフィルターを定義する `ItemHasRegularExpressionMatch` ルール要素の名前です。</span><span class="sxs-lookup"><span data-stu-id="85619-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-863">要件</span><span class="sxs-lookup"><span data-stu-id="85619-863">Requirements</span></span>

|<span data-ttu-id="85619-864">要件</span><span class="sxs-lookup"><span data-stu-id="85619-864">Requirement</span></span>| <span data-ttu-id="85619-865">値</span><span class="sxs-lookup"><span data-stu-id="85619-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-866">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-866">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-867">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-867">1.0</span></span>|
|[<span data-ttu-id="85619-868">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-869">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-869">ReadItem</span></span>|
|[<span data-ttu-id="85619-870">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-871">閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-872">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-872">Returns:</span></span>

<span data-ttu-id="85619-873">マニフェスト XML ファイルで定義された正規表現に一致する文字列を含む配列です。</span><span class="sxs-lookup"><span data-stu-id="85619-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="85619-874">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="85619-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="85619-875">配列. < 文字列 ></span><span class="sxs-lookup"><span data-stu-id="85619-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="85619-876">例</span><span class="sxs-lookup"><span data-stu-id="85619-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="85619-877">[選択された非同期データを取得（getSelectedDataAsync）](coercionType、[オプション]、コールバック) → {文字列}</span><span class="sxs-lookup"><span data-stu-id="85619-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="85619-878">メッセージの件名または本文から非同期的に選択したデータを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="85619-p158">選択したデータがなく、カーソルが本文または件名内にある場合、メソッドは選択したデータに対して Null を返します。本文または件名以外のフィールドが選択されている場合、メソッドは `InvalidSelection` エラーを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-881">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-881">Parameters:</span></span>

|<span data-ttu-id="85619-882">名前</span><span class="sxs-lookup"><span data-stu-id="85619-882">Name</span></span>| <span data-ttu-id="85619-883">種類</span><span class="sxs-lookup"><span data-stu-id="85619-883">Type</span></span>| <span data-ttu-id="85619-884">特性</span><span class="sxs-lookup"><span data-stu-id="85619-884">Attributes</span></span>| <span data-ttu-id="85619-885">記述</span><span class="sxs-lookup"><span data-stu-id="85619-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="85619-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="85619-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="85619-p159">データの形式を要求します。テキストの場合、メソッドは文字列としてプレーン テキストを返し、存在する HTML タグは全て削除されます。HTMLの場合、メソッドは、プレーンテキストまたは HTML のいずれの場合も選択されたテキストを返します。</span><span class="sxs-lookup"><span data-stu-id="85619-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="85619-890">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-890">Object</span></span>| <span data-ttu-id="85619-891">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-891">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-892">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-893">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-893">Object</span></span>| <span data-ttu-id="85619-894">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-894">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-895">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="85619-896">関数</span><span class="sxs-lookup"><span data-stu-id="85619-896">function</span></span>||<span data-ttu-id="85619-897">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="85619-898">コールバック メソッドから選択したデータへアクセスするには、`asyncResult.value.data`を呼び出してください。</span><span class="sxs-lookup"><span data-stu-id="85619-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="85619-899">選択元のソース プロパティにアクセスするには、 `asyncResult.value.sourceProperty` を呼び出します。これは `body`   または `subject`  になります。</span><span class="sxs-lookup"><span data-stu-id="85619-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-900">要件</span><span class="sxs-lookup"><span data-stu-id="85619-900">Requirements</span></span>

|<span data-ttu-id="85619-901">要件</span><span class="sxs-lookup"><span data-stu-id="85619-901">Requirement</span></span>| <span data-ttu-id="85619-902">値</span><span class="sxs-lookup"><span data-stu-id="85619-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-903">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-903">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-904">1.2</span><span class="sxs-lookup"><span data-stu-id="85619-904">1.2</span></span>|
|[<span data-ttu-id="85619-905">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-906">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-907">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-908">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="85619-909">次を返します :</span><span class="sxs-lookup"><span data-stu-id="85619-909">Returns:</span></span>

<span data-ttu-id="85619-910">`coercionType`に決定された書式設定の文字列として選択されたデータです。</span><span class="sxs-lookup"><span data-stu-id="85619-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="85619-911">

<dt>種類</dt>

</span><span class="sxs-lookup"><span data-stu-id="85619-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="85619-912">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="85619-913">例</span><span class="sxs-lookup"><span data-stu-id="85619-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="85619-914">[非同期敵にカスタムプロパティをロード（loadCustomPropertiesAsync）](コールバック、[ユーザコンテキスト（userContext）])</span><span class="sxs-lookup"><span data-stu-id="85619-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="85619-915">選択された項目でこのアドインのカスタム プロパティを非同期に読み込みます。</span><span class="sxs-lookup"><span data-stu-id="85619-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="85619-p161">カスタム プロパティは、アプリケーションごと、項目ごとにキーと値のペアとして保管されます。このメソッドは、コールバックで `CustomProperties` オブジェクトを返します。このオブジェクトは、現在の項目および現在のアドインに固有のカスタム プロパティにアクセスするためのメソッドを提供します。カスタム プロパティは、項目上では暗号化されていません。そのため、セキュリティ保護記憶域として使用するべきではありません。</span><span class="sxs-lookup"><span data-stu-id="85619-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-919">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-919">Parameters:</span></span>

|<span data-ttu-id="85619-920">名前</span><span class="sxs-lookup"><span data-stu-id="85619-920">Name</span></span>| <span data-ttu-id="85619-921">種類</span><span class="sxs-lookup"><span data-stu-id="85619-921">Type</span></span>| <span data-ttu-id="85619-922">特性</span><span class="sxs-lookup"><span data-stu-id="85619-922">Attributes</span></span>| <span data-ttu-id="85619-923">記述</span><span class="sxs-lookup"><span data-stu-id="85619-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="85619-924">関数</span><span class="sxs-lookup"><span data-stu-id="85619-924">function</span></span>||<span data-ttu-id="85619-925">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="85619-926">カスタム プロパティは `asyncResult.value` プロパティの [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) オブジェクトとして提供されます。</span><span class="sxs-lookup"><span data-stu-id="85619-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="85619-927">項目からカスタム プロパティを取得、設定、削除して、サーバーにカスタム プロパティのセット バックに対する変更を保存するのに、このオブジェクトを使用できます。</span><span class="sxs-lookup"><span data-stu-id="85619-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="85619-928">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-928">Object</span></span>| <span data-ttu-id="85619-929">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-929">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-930">開発者は、コールバック関数でアクセスしたいオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="85619-931">このオブジェクトには、コールバック関数の `asyncResult.asyncContext` プロパティによってアクセスすることができます。</span><span class="sxs-lookup"><span data-stu-id="85619-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-932">要件</span><span class="sxs-lookup"><span data-stu-id="85619-932">Requirements</span></span>

|<span data-ttu-id="85619-933">要件</span><span class="sxs-lookup"><span data-stu-id="85619-933">Requirement</span></span>| <span data-ttu-id="85619-934">値</span><span class="sxs-lookup"><span data-stu-id="85619-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-935">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-936">1.0</span><span class="sxs-lookup"><span data-stu-id="85619-936">1.0</span></span>|
|[<span data-ttu-id="85619-937">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-938">項目を閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-938">ReadItem</span></span>|
|[<span data-ttu-id="85619-939">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-940">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="85619-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-941">例</span><span class="sxs-lookup"><span data-stu-id="85619-941">Example</span></span>

<span data-ttu-id="85619-p164">次のコード例では、`loadCustomPropertiesAsync` メソッドを使用して、現在のアイテムに固有のカスタム プロパティを非同期的に読み込む方法を示します。また、`CustomProperties.saveAsync` メソッドを使用して、これらのプロパティをサーバーに保存する方法も紹介します。カスタム プロパティをロードした後、このコード サンプルでは `CustomProperties.get` メソッドを使用してカスタム プロパティ `myProp` を読み取り、`CustomProperties.set` メソッドでカスタム プロパティ `otherProp` を書き込み、最後に `saveAsync` メソッドを呼び出して、カスタム プロパティを保存します。</span><span class="sxs-lookup"><span data-stu-id="85619-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="85619-945">[非同期敵に添付ファイルを削除（removeAttachmentAsync）]([添付ファイルID（attachmentId）]、[オプション]、 [コールバック])</span><span class="sxs-lookup"><span data-stu-id="85619-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="85619-946">メッセージまたはアポイントメントから添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="85619-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="85619-p165">`removeAttachmentAsync`メソッドは、指定した識別子の添付ファイルを項目 から削除します。ベスト プラクティスとして、同じメール アプリが同じセッションで添付ファイルを追加した場合にのみ、その添付ファイルの識別子を使用して添付ファイルを削除することをお勧めします。Outlook Web アプリ とデバイス用 OWA では、添付ファイルの識別子は同じセッション内でのみ有効です。ユーザーがアプリを閉じるか、ユーザーがインライン フォームで新規作成を開始してインライン フォームが表示され、別のウィンドウで操作を継続すると、セッションは終了します。</span><span class="sxs-lookup"><span data-stu-id="85619-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-951">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-951">Parameters:</span></span>

|<span data-ttu-id="85619-952">名前</span><span class="sxs-lookup"><span data-stu-id="85619-952">Name</span></span>| <span data-ttu-id="85619-953">種類</span><span class="sxs-lookup"><span data-stu-id="85619-953">Type</span></span>| <span data-ttu-id="85619-954">特性</span><span class="sxs-lookup"><span data-stu-id="85619-954">Attributes</span></span>| <span data-ttu-id="85619-955">記述</span><span class="sxs-lookup"><span data-stu-id="85619-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="85619-956">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-956">String</span></span>||<span data-ttu-id="85619-p166">削除する添付ファイルの識別子です。最大文字数は 100 文字です。</span><span class="sxs-lookup"><span data-stu-id="85619-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="85619-959">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-959">Object</span></span>| <span data-ttu-id="85619-960">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-960">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-961">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-962">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-962">Object</span></span>| <span data-ttu-id="85619-963">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-963">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-964">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="85619-965">関数</span><span class="sxs-lookup"><span data-stu-id="85619-965">function</span></span>| <span data-ttu-id="85619-966">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-966">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-967">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="85619-968">添付ファイルの削除に失敗すると、`asyncResult.error` プロパティにはエラー コードとエラーの理由が含まれます。</span><span class="sxs-lookup"><span data-stu-id="85619-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="85619-969">エラー</span><span class="sxs-lookup"><span data-stu-id="85619-969">Errors</span></span>

| <span data-ttu-id="85619-970">エラー コード</span><span class="sxs-lookup"><span data-stu-id="85619-970">Error code</span></span> | <span data-ttu-id="85619-971">記述</span><span class="sxs-lookup"><span data-stu-id="85619-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="85619-972">添付ファイルの識別子が存在しません。</span><span class="sxs-lookup"><span data-stu-id="85619-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-973">要件</span><span class="sxs-lookup"><span data-stu-id="85619-973">Requirements</span></span>

|<span data-ttu-id="85619-974">要件</span><span class="sxs-lookup"><span data-stu-id="85619-974">Requirement</span></span>| <span data-ttu-id="85619-975">値</span><span class="sxs-lookup"><span data-stu-id="85619-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-976">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-976">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-977">1.1</span><span class="sxs-lookup"><span data-stu-id="85619-977">1.1</span></span>|
|[<span data-ttu-id="85619-978">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-979">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-980">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-981">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-982">例</span><span class="sxs-lookup"><span data-stu-id="85619-982">Example</span></span>

<span data-ttu-id="85619-983">次のコードは、'0' の識別子を持つ添付ファイルを削除します。</span><span class="sxs-lookup"><span data-stu-id="85619-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="85619-984">[非同期保存（saveAsync）] ([オプション] 、コールバック)</span><span class="sxs-lookup"><span data-stu-id="85619-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="85619-985">アイテムを非同期的に保存します。</span><span class="sxs-lookup"><span data-stu-id="85619-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="85619-p167">呼び出されると、このメソッドは現在のメッセージを下書きとして保存し、コールバック メソッド経由で項目 ID を返します。Outlook Web App またはオンライン モードの Outlook では、サーバーに項目が保存されます。キャッシュ モードの Outlook では、ローカル キャッシュに項目が保存されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-989">アドインが、EWS または REST API を使用しようとして`itemId`を取得するために、新規作成モードでアイテム上の`saveAsync`を呼び出す場合、Outlook キャッシュ モードでは、項目がサーバーと実際に同期するまでに時間がかかる場合があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="85619-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="85619-990">項目が同期されるまで、 `itemId` を使用すると、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="85619-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="85619-p169">アポイントメントはドラフト状態にはならないため、作成モードでアポイントメントに`saveAsync`が呼び出される場合、その項目はユーザーのカレンダーに通常のアポイントメントとして保存されます。以前に保存されていない新しいアポイントメントの場合、招待状は送信されません。既存のアポイントメントを保存すると、追加または削除された出席者に更新が送信されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="85619-994">次のクライアントは、新規作成モードでアポイントメント上の `saveAsync` に対して様々な行動を行ないます：</span><span class="sxs-lookup"><span data-stu-id="85619-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="85619-995">Mac Outlook は、作成モードの会議で`saveAsync`をサポートしていません。</span><span class="sxs-lookup"><span data-stu-id="85619-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="85619-996">Mac Outlookの会議場で  `saveAsync` を呼びだすと、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="85619-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="85619-997">作成モードのアポイントメント上で`saveAsync`が呼び出されると、Outlook on the web は常に、招待または更新を送信します。</span><span class="sxs-lookup"><span data-stu-id="85619-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-998">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-998">Parameters:</span></span>

|<span data-ttu-id="85619-999">名前</span><span class="sxs-lookup"><span data-stu-id="85619-999">Name</span></span>| <span data-ttu-id="85619-1000">種類</span><span class="sxs-lookup"><span data-stu-id="85619-1000">Type</span></span>| <span data-ttu-id="85619-1001">特性</span><span class="sxs-lookup"><span data-stu-id="85619-1001">Attributes</span></span>| <span data-ttu-id="85619-1002">記述</span><span class="sxs-lookup"><span data-stu-id="85619-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="85619-1003">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-1003">Object</span></span>| <span data-ttu-id="85619-1004">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-1005">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-1006">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-1006">Object</span></span>| <span data-ttu-id="85619-1007">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-1008">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="85619-1009">関数</span><span class="sxs-lookup"><span data-stu-id="85619-1009">function</span></span>||<span data-ttu-id="85619-1010">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="85619-1011">成功すると、項目の識別子が`asyncResult.value`プロパティに提供されます。</span><span class="sxs-lookup"><span data-stu-id="85619-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="85619-1012">要件</span><span class="sxs-lookup"><span data-stu-id="85619-1012">Requirements</span></span>

|<span data-ttu-id="85619-1013">要件</span><span class="sxs-lookup"><span data-stu-id="85619-1013">Requirement</span></span>| <span data-ttu-id="85619-1014">値</span><span class="sxs-lookup"><span data-stu-id="85619-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-1015">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-1015">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="85619-1016">1.3</span></span>|
|[<span data-ttu-id="85619-1017">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-1018">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-1019">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-1020">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="85619-1021">例</span><span class="sxs-lookup"><span data-stu-id="85619-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="85619-p171">次の例は、コールバック関数に渡される `result` パラメータの例です。`value` プロパティには、項目の項目 ID が含まれます。</span><span class="sxs-lookup"><span data-stu-id="85619-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="85619-1024">setSelectedDataAsync(データ, [オプション], コールバック)</span><span class="sxs-lookup"><span data-stu-id="85619-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="85619-1025">メッセージの本文または件名に非同期的にデータを挿入します。</span><span class="sxs-lookup"><span data-stu-id="85619-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="85619-p172">`setSelectedDataAsync`メソッドは、指定された文字列を項目の件名または本文のカーソル位置に挿入する、または、エディターでテキストが選択されている場合は、選択されたテキストを書き換えます。本文または件名フィールド内にカーソルがない場合は、エラーが返されます。挿入後、カーソルは挿入されたコンテンツの末尾に置かれます。</span><span class="sxs-lookup"><span data-stu-id="85619-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="85619-1029">パラメータ：</span><span class="sxs-lookup"><span data-stu-id="85619-1029">Parameters:</span></span>

|<span data-ttu-id="85619-1030">名前</span><span class="sxs-lookup"><span data-stu-id="85619-1030">Name</span></span>| <span data-ttu-id="85619-1031">種類</span><span class="sxs-lookup"><span data-stu-id="85619-1031">Type</span></span>| <span data-ttu-id="85619-1032">特性</span><span class="sxs-lookup"><span data-stu-id="85619-1032">Attributes</span></span>| <span data-ttu-id="85619-1033">記述</span><span class="sxs-lookup"><span data-stu-id="85619-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="85619-1034">文字列</span><span class="sxs-lookup"><span data-stu-id="85619-1034">String</span></span>||<span data-ttu-id="85619-p173">挿入されるデータです。データの長さは 1,000,000 文字以内でなければなりません。1,000,000 文字を超えるデータが渡されると、 `ArgumentOutOfRange` の例外がスローされます。</span><span class="sxs-lookup"><span data-stu-id="85619-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="85619-1038">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-1038">Object</span></span>| <span data-ttu-id="85619-1039">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-1040">次のプロパティを 1 つ以上含むオブジェクト リテラルです。</span><span class="sxs-lookup"><span data-stu-id="85619-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="85619-1041">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="85619-1041">Object</span></span>| <span data-ttu-id="85619-1042">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-1043">開発者は、コールバック メソッドでアクセスしたい任意のオブジェクトを提供できます。</span><span class="sxs-lookup"><span data-stu-id="85619-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="85619-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="85619-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="85619-1045">&lt;任意&gt;</span><span class="sxs-lookup"><span data-stu-id="85619-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="85619-p174">`text` の場合、Outlook Web アプリ と Outlook で現在のスタイルが適用されます。フィールドが HTML エディターの場合、データが HTML の場合でも、テキスト データのみが挿入されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="85619-p175">`html` とフィールドが HTML をサポートする場合 (件名はサポートしない)、Office Web App では現在のスタイルが適用され、Outlook では既定のスタイルが適用されます。フィールドがテキスト フィールドの場合、`InvalidDataFormat` エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="85619-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="85619-1050">`coercionType` が設定されていない場合、結果はフィールドによって変わります：フィールドが HTML の場合は HTML が使用されます。フィールドがテキストの場合はプレーン テキストが使用されます。</span><span class="sxs-lookup"><span data-stu-id="85619-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="85619-1051">関数</span><span class="sxs-lookup"><span data-stu-id="85619-1051">function</span></span>||<span data-ttu-id="85619-1052">メソッドが完了すると、`callback` パラメータで渡された関数は単一パラメータ `asyncResult` で呼び出されます。このパラメータは、[`AsyncResult`](/javascript/api/office/office.asyncresult) オブジェクトです。</span><span class="sxs-lookup"><span data-stu-id="85619-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="85619-1053">要件</span><span class="sxs-lookup"><span data-stu-id="85619-1053">Requirements</span></span>

|<span data-ttu-id="85619-1054">要件</span><span class="sxs-lookup"><span data-stu-id="85619-1054">Requirement</span></span>| <span data-ttu-id="85619-1055">値</span><span class="sxs-lookup"><span data-stu-id="85619-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="85619-1056">​最小限のメールボックス要件セットバージョン</span><span class="sxs-lookup"><span data-stu-id="85619-1056">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="85619-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="85619-1057">1.2</span></span>|
|[<span data-ttu-id="85619-1058">​最小限のアクセス許可レベル​</span><span class="sxs-lookup"><span data-stu-id="85619-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="85619-1059">[項目の読み取り/書き込み（ReadWriteItem）]</span><span class="sxs-lookup"><span data-stu-id="85619-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="85619-1060">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="85619-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="85619-1061">新規作成</span><span class="sxs-lookup"><span data-stu-id="85619-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="85619-1062">例</span><span class="sxs-lookup"><span data-stu-id="85619-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```