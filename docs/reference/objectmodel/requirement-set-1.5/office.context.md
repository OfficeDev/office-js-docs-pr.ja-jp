# <a name="context"></a><span data-ttu-id="16933-101">コンテキスト</span><span class="sxs-lookup"><span data-stu-id="16933-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="16933-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="16933-102">[Office](Office.md).context</span></span>

<span data-ttu-id="16933-p101">Office.context の名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧では、Outlook アドインで使うインターフェイスのみを記載しています。Office.context の名前空間の完全な一覧については、[ 共有 API の 中のOffice.context リファレンス](/javascript/api/office/office.context) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="16933-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="16933-105">要件</span><span class="sxs-lookup"><span data-stu-id="16933-105">Requirements</span></span>

|<span data-ttu-id="16933-106">要件</span><span class="sxs-lookup"><span data-stu-id="16933-106">Requirement</span></span>| <span data-ttu-id="16933-107">値</span><span class="sxs-lookup"><span data-stu-id="16933-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="16933-108">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="16933-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16933-109">1.0</span><span class="sxs-lookup"><span data-stu-id="16933-109">1.0</span></span>|
|[<span data-ttu-id="16933-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16933-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16933-111">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16933-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="16933-112">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="16933-112">Members and methods</span></span>

| <span data-ttu-id="16933-113">メンバー</span><span class="sxs-lookup"><span data-stu-id="16933-113">Member</span></span> | <span data-ttu-id="16933-114">種類</span><span class="sxs-lookup"><span data-stu-id="16933-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="16933-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="16933-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="16933-116">メンバー</span><span class="sxs-lookup"><span data-stu-id="16933-116">Member</span></span> |
| [<span data-ttu-id="16933-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="16933-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="16933-118">メンバー</span><span class="sxs-lookup"><span data-stu-id="16933-118">Member</span></span> |
| [<span data-ttu-id="16933-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="16933-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings) | <span data-ttu-id="16933-120">メンバー</span><span class="sxs-lookup"><span data-stu-id="16933-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="16933-121">名前空間</span><span class="sxs-lookup"><span data-stu-id="16933-121">Namespaces</span></span>

<span data-ttu-id="16933-122">[メ―ルボックス](office.context.mailbox.md): Microsoft Outlook と Microsoft Outlook on the Web の Outlook アドイン オブジェクト モデルへアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="16933-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="16933-123">メンバー</span><span class="sxs-lookup"><span data-stu-id="16933-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="16933-124">displayLanguage: 文字列</span><span class="sxs-lookup"><span data-stu-id="16933-124">displayLanguage :String</span></span>

<span data-ttu-id="16933-125">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="16933-126">値`displayLanguage`は電流を反映する**言語を表示する** Office ホスト アプリケーション内で、**ファイル > オプション > 言語**によって指定された設定</span><span class="sxs-lookup"><span data-stu-id="16933-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="16933-127">種類:</span><span class="sxs-lookup"><span data-stu-id="16933-127">Type:</span></span>

*   <span data-ttu-id="16933-128">文字列</span><span class="sxs-lookup"><span data-stu-id="16933-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="16933-129">要件</span><span class="sxs-lookup"><span data-stu-id="16933-129">Requirements</span></span>

|<span data-ttu-id="16933-130">要件</span><span class="sxs-lookup"><span data-stu-id="16933-130">Requirement</span></span>| <span data-ttu-id="16933-131">値</span><span class="sxs-lookup"><span data-stu-id="16933-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="16933-132">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="16933-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16933-133">1.0</span><span class="sxs-lookup"><span data-stu-id="16933-133">1.0</span></span>|
|[<span data-ttu-id="16933-134">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16933-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16933-135">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16933-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="16933-136">例</span><span class="sxs-lookup"><span data-stu-id="16933-136">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a><span data-ttu-id="16933-137">officeTheme: オブジェクト</span><span class="sxs-lookup"><span data-stu-id="16933-137">officeTheme :Object</span></span>

<span data-ttu-id="16933-138">Office のテーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="16933-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="16933-139">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="16933-139">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="16933-p102">Office のテーマの色を使うと、**ファイル >Office アカウント >Office テーマ UI**によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Officeの テーマの色は、メール アドインと作業ウィンドウ アドインに適合しています。</span><span class="sxs-lookup"><span data-stu-id="16933-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="16933-142">種類:</span><span class="sxs-lookup"><span data-stu-id="16933-142">Type:</span></span>

*   <span data-ttu-id="16933-143">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="16933-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="16933-144">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="16933-144">Properties:</span></span>

|<span data-ttu-id="16933-145">名前</span><span class="sxs-lookup"><span data-stu-id="16933-145">Name</span></span>| <span data-ttu-id="16933-146">種類</span><span class="sxs-lookup"><span data-stu-id="16933-146">Type</span></span>| <span data-ttu-id="16933-147">説明</span><span class="sxs-lookup"><span data-stu-id="16933-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="16933-148">文字列</span><span class="sxs-lookup"><span data-stu-id="16933-148">String</span></span>|<span data-ttu-id="16933-149">Office テーマの本文背景色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="16933-150">文字列</span><span class="sxs-lookup"><span data-stu-id="16933-150">String</span></span>|<span data-ttu-id="16933-151">Office テーマの本文の前景色を 16 進トリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="16933-152">文字列</span><span class="sxs-lookup"><span data-stu-id="16933-152">String</span></span>|<span data-ttu-id="16933-153">Office テーマコントロールの背景色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="16933-154">文字列</span><span class="sxs-lookup"><span data-stu-id="16933-154">String</span></span>|<span data-ttu-id="16933-155">Office テーマの本文コントロール色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="16933-156">要件</span><span class="sxs-lookup"><span data-stu-id="16933-156">Requirements</span></span>

|<span data-ttu-id="16933-157">要件</span><span class="sxs-lookup"><span data-stu-id="16933-157">Requirement</span></span>| <span data-ttu-id="16933-158">値</span><span class="sxs-lookup"><span data-stu-id="16933-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="16933-159">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="16933-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16933-160">1.3</span><span class="sxs-lookup"><span data-stu-id="16933-160">1.3</span></span>|
|[<span data-ttu-id="16933-161">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16933-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16933-162">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16933-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="16933-163">例</span><span class="sxs-lookup"><span data-stu-id="16933-163">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings"></a><span data-ttu-id="16933-164">roamingSettings:[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="16933-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span></span>

<span data-ttu-id="16933-165">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="16933-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="16933-166">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのためのデータの保存やアクセスができます。そのため、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションからメール アドインを実行してもこのデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="16933-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="16933-167">種類:</span><span class="sxs-lookup"><span data-stu-id="16933-167">Type:</span></span>

*   [<span data-ttu-id="16933-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="16933-168">RoamingSettings</span></span>](/javascript/api/outlook_1_5/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="16933-169">要件</span><span class="sxs-lookup"><span data-stu-id="16933-169">Requirements</span></span>

|<span data-ttu-id="16933-170">要件</span><span class="sxs-lookup"><span data-stu-id="16933-170">Requirement</span></span>| <span data-ttu-id="16933-171">値</span><span class="sxs-lookup"><span data-stu-id="16933-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="16933-172">メールボックス要件の最小バージョン</span><span class="sxs-lookup"><span data-stu-id="16933-172">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="16933-173">1.0</span><span class="sxs-lookup"><span data-stu-id="16933-173">1.0</span></span>|
|[<span data-ttu-id="16933-174">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="16933-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="16933-175">制限あり</span><span class="sxs-lookup"><span data-stu-id="16933-175">Restricted</span></span>|
|[<span data-ttu-id="16933-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="16933-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="16933-177">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="16933-177">Compose or read</span></span>|