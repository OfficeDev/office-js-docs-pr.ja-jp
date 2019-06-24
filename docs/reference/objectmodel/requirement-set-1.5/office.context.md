---
title: Office.context - 要件セット 1.5
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 864fd8a92d30526cfb87c76754934fd9a9cbd827
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127318"
---
# <a name="context"></a><span data-ttu-id="08a1f-102">context</span><span class="sxs-lookup"><span data-stu-id="08a1f-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="08a1f-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="08a1f-103">[Office](Office.md).context</span></span>

<span data-ttu-id="08a1f-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="08a1f-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="08a1f-106">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-106">Requirements</span></span>

|<span data-ttu-id="08a1f-107">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-107">Requirement</span></span>| <span data-ttu-id="08a1f-108">値</span><span class="sxs-lookup"><span data-stu-id="08a1f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="08a1f-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08a1f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="08a1f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="08a1f-110">1.0</span></span>|
|[<span data-ttu-id="08a1f-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08a1f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08a1f-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08a1f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="08a1f-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="08a1f-113">Members and methods</span></span>

| <span data-ttu-id="08a1f-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="08a1f-114">Member</span></span> | <span data-ttu-id="08a1f-115">種類</span><span class="sxs-lookup"><span data-stu-id="08a1f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="08a1f-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="08a1f-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="08a1f-117">Member</span><span class="sxs-lookup"><span data-stu-id="08a1f-117">Member</span></span> |
| [<span data-ttu-id="08a1f-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="08a1f-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="08a1f-119">Member</span><span class="sxs-lookup"><span data-stu-id="08a1f-119">Member</span></span> |
| [<span data-ttu-id="08a1f-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="08a1f-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="08a1f-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="08a1f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="08a1f-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="08a1f-122">Namespaces</span></span>

<span data-ttu-id="08a1f-123">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="08a1f-124">Members</span><span class="sxs-lookup"><span data-stu-id="08a1f-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="08a1f-125">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="08a1f-125">displayLanguage: String</span></span>

<span data-ttu-id="08a1f-126">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="08a1f-127">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="08a1f-128">型</span><span class="sxs-lookup"><span data-stu-id="08a1f-128">Type</span></span>

*   <span data-ttu-id="08a1f-129">String</span><span class="sxs-lookup"><span data-stu-id="08a1f-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="08a1f-130">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-130">Requirements</span></span>

|<span data-ttu-id="08a1f-131">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-131">Requirement</span></span>| <span data-ttu-id="08a1f-132">値</span><span class="sxs-lookup"><span data-stu-id="08a1f-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="08a1f-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08a1f-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="08a1f-134">1.0</span><span class="sxs-lookup"><span data-stu-id="08a1f-134">1.0</span></span>|
|[<span data-ttu-id="08a1f-135">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08a1f-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08a1f-136">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08a1f-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="08a1f-137">例</span><span class="sxs-lookup"><span data-stu-id="08a1f-137">Example</span></span>

```javascript
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

#### <a name="officetheme-object"></a><span data-ttu-id="08a1f-138">officeTheme: オブジェクト</span><span class="sxs-lookup"><span data-stu-id="08a1f-138">officeTheme: Object</span></span>

<span data-ttu-id="08a1f-139">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="08a1f-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="08a1f-140">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="08a1f-140">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="08a1f-p102">Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。</span><span class="sxs-lookup"><span data-stu-id="08a1f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="08a1f-143">型</span><span class="sxs-lookup"><span data-stu-id="08a1f-143">Type</span></span>

*   <span data-ttu-id="08a1f-144">Object</span><span class="sxs-lookup"><span data-stu-id="08a1f-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="08a1f-145">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="08a1f-145">Properties:</span></span>

|<span data-ttu-id="08a1f-146">名前</span><span class="sxs-lookup"><span data-stu-id="08a1f-146">Name</span></span>| <span data-ttu-id="08a1f-147">種類</span><span class="sxs-lookup"><span data-stu-id="08a1f-147">Type</span></span>| <span data-ttu-id="08a1f-148">説明</span><span class="sxs-lookup"><span data-stu-id="08a1f-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="08a1f-149">String</span><span class="sxs-lookup"><span data-stu-id="08a1f-149">String</span></span>|<span data-ttu-id="08a1f-150">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="08a1f-151">String</span><span class="sxs-lookup"><span data-stu-id="08a1f-151">String</span></span>|<span data-ttu-id="08a1f-152">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="08a1f-153">String</span><span class="sxs-lookup"><span data-stu-id="08a1f-153">String</span></span>|<span data-ttu-id="08a1f-154">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="08a1f-155">String</span><span class="sxs-lookup"><span data-stu-id="08a1f-155">String</span></span>|<span data-ttu-id="08a1f-156">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="08a1f-157">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-157">Requirements</span></span>

|<span data-ttu-id="08a1f-158">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-158">Requirement</span></span>| <span data-ttu-id="08a1f-159">値</span><span class="sxs-lookup"><span data-stu-id="08a1f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="08a1f-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08a1f-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="08a1f-161">1.3</span><span class="sxs-lookup"><span data-stu-id="08a1f-161">1.3</span></span>|
|[<span data-ttu-id="08a1f-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08a1f-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08a1f-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08a1f-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="08a1f-164">例</span><span class="sxs-lookup"><span data-stu-id="08a1f-164">Example</span></span>

```javascript
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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings"></a><span data-ttu-id="08a1f-165">roamingSettings: [roamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="08a1f-165">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span></span>

<span data-ttu-id="08a1f-166">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="08a1f-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="08a1f-167">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="08a1f-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="08a1f-168">型</span><span class="sxs-lookup"><span data-stu-id="08a1f-168">Type</span></span>

*   [<span data-ttu-id="08a1f-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="08a1f-169">RoamingSettings</span></span>](/javascript/api/outlook_1_5/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="08a1f-170">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-170">Requirements</span></span>

|<span data-ttu-id="08a1f-171">要件</span><span class="sxs-lookup"><span data-stu-id="08a1f-171">Requirement</span></span>| <span data-ttu-id="08a1f-172">値</span><span class="sxs-lookup"><span data-stu-id="08a1f-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="08a1f-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="08a1f-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="08a1f-174">1.0</span><span class="sxs-lookup"><span data-stu-id="08a1f-174">1.0</span></span>|
|[<span data-ttu-id="08a1f-175">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="08a1f-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="08a1f-176">制限あり</span><span class="sxs-lookup"><span data-stu-id="08a1f-176">Restricted</span></span>|
|[<span data-ttu-id="08a1f-177">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="08a1f-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="08a1f-178">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="08a1f-178">Compose or Read</span></span>|
