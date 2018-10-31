
# <a name="context"></a><span data-ttu-id="8546e-101">コンテキスト</span><span class="sxs-lookup"><span data-stu-id="8546e-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="8546e-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="8546e-102">[Office](Office.md).context</span></span>

<span data-ttu-id="8546e-p101">Office.context の名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧では、Outlook アドインで使うインターフェイスのみを記載しています。Office.context の名前空間の完全な一覧については、[ 共有 API の 中のOffice.context リファレンス](/javascript/api/office/office.context) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8546e-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8546e-105">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-105">Requirements</span></span>

|<span data-ttu-id="8546e-106">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-106">Requirement</span></span>| <span data-ttu-id="8546e-107">値</span><span class="sxs-lookup"><span data-stu-id="8546e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8546e-108">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8546e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8546e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8546e-109">1.0</span></span>|
|[<span data-ttu-id="8546e-110">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8546e-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8546e-111">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8546e-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="8546e-112">名前空間</span><span class="sxs-lookup"><span data-stu-id="8546e-112">Namespaces</span></span>

<span data-ttu-id="8546e-113">[メ―ルボックス](office.context.mailbox.md): Microsoft Outlook と Microsoft Outlook on the Web の Outlook アドイン オブジェクト モデルへアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="8546e-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="8546e-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="8546e-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="8546e-115">displayLanguage: 文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-115">displayLanguage :String</span></span>

<span data-ttu-id="8546e-116">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="8546e-117">値`displayLanguage`は電流を反映する**言語を表示する** Office ホスト アプリケーション内で、**ファイル > オプション > 言語**によって指定された設定</span><span class="sxs-lookup"><span data-stu-id="8546e-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="8546e-118">種類:</span><span class="sxs-lookup"><span data-stu-id="8546e-118">Type:</span></span>

*   <span data-ttu-id="8546e-119">文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8546e-120">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-120">Requirements</span></span>

|<span data-ttu-id="8546e-121">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-121">Requirement</span></span>| <span data-ttu-id="8546e-122">値</span><span class="sxs-lookup"><span data-stu-id="8546e-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="8546e-123">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8546e-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8546e-124">1.0</span><span class="sxs-lookup"><span data-stu-id="8546e-124">1.0</span></span>|
|[<span data-ttu-id="8546e-125">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8546e-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8546e-126">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8546e-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8546e-127">例</span><span class="sxs-lookup"><span data-stu-id="8546e-127">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="8546e-128">officeTheme: オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8546e-128">officeTheme :Object</span></span>

<span data-ttu-id="8546e-129">Office のテーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="8546e-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="8546e-130">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8546e-130">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8546e-p102">Office のテーマの色を使うと、**ファイル >Office アカウント >Office テーマ UI**によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Officeの テーマの色は、メール アドインと作業ウィンドウ アドインに適合しています。</span><span class="sxs-lookup"><span data-stu-id="8546e-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8546e-133">型:</span><span class="sxs-lookup"><span data-stu-id="8546e-133">Type:</span></span>

*   <span data-ttu-id="8546e-134">オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8546e-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="8546e-135">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8546e-135">Properties:</span></span>

|<span data-ttu-id="8546e-136">名前</span><span class="sxs-lookup"><span data-stu-id="8546e-136">Name</span></span>| <span data-ttu-id="8546e-137">種類</span><span class="sxs-lookup"><span data-stu-id="8546e-137">Type</span></span>| <span data-ttu-id="8546e-138">説明</span><span class="sxs-lookup"><span data-stu-id="8546e-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="8546e-139">文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-139">String</span></span>|<span data-ttu-id="8546e-140">Office テーマの本文背景色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="8546e-141">文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-141">String</span></span>|<span data-ttu-id="8546e-142">Office テーマの本文の前景色を 16 進トリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="8546e-143">文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-143">String</span></span>|<span data-ttu-id="8546e-144">Office テーマコントロールの背景色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="8546e-145">文字列</span><span class="sxs-lookup"><span data-stu-id="8546e-145">String</span></span>|<span data-ttu-id="8546e-146">Office テーマの本文コントロール色を 16 進法のカラートリプレットとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8546e-147">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-147">Requirements</span></span>

|<span data-ttu-id="8546e-148">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-148">Requirement</span></span>| <span data-ttu-id="8546e-149">値</span><span class="sxs-lookup"><span data-stu-id="8546e-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="8546e-150">メールボックス要件セットの最小バージョン</span><span class="sxs-lookup"><span data-stu-id="8546e-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8546e-151">1.3</span><span class="sxs-lookup"><span data-stu-id="8546e-151">1.3</span></span>|
|[<span data-ttu-id="8546e-152">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8546e-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8546e-153">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8546e-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8546e-154">例</span><span class="sxs-lookup"><span data-stu-id="8546e-154">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="8546e-155">roamingSettings:[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="8546e-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="8546e-156">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8546e-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8546e-157">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのためのデータの保存やアクセスができます。そのため、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションからメール アドインを実行してもこのデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="8546e-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8546e-158">種類:</span><span class="sxs-lookup"><span data-stu-id="8546e-158">Type:</span></span>

*   [<span data-ttu-id="8546e-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8546e-159">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8546e-160">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-160">Requirements</span></span>

|<span data-ttu-id="8546e-161">要件</span><span class="sxs-lookup"><span data-stu-id="8546e-161">Requirement</span></span>| <span data-ttu-id="8546e-162">値</span><span class="sxs-lookup"><span data-stu-id="8546e-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="8546e-163">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8546e-163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8546e-164">1.0</span><span class="sxs-lookup"><span data-stu-id="8546e-164">1.0</span></span>|
|[<span data-ttu-id="8546e-165">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8546e-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8546e-166">制限あり</span><span class="sxs-lookup"><span data-stu-id="8546e-166">Restricted</span></span>|
|[<span data-ttu-id="8546e-167">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8546e-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8546e-168">作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="8546e-168">Compose or read</span></span>|