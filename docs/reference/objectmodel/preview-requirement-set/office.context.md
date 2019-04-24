---
title: Office コンテキスト-プレビュー要件セット
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: a1e01142a4c0b84a4afcba89f76766d28595ba95
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450346"
---
# <a name="context"></a><span data-ttu-id="c9afc-102">context</span><span class="sxs-lookup"><span data-stu-id="c9afc-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="c9afc-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="c9afc-103">[Office](Office.md).context</span></span>

<span data-ttu-id="c9afc-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c9afc-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9afc-106">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-106">Requirements</span></span>

|<span data-ttu-id="c9afc-107">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-107">Requirement</span></span>| <span data-ttu-id="c9afc-108">値</span><span class="sxs-lookup"><span data-stu-id="c9afc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9afc-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9afc-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9afc-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c9afc-110">1.0</span></span>|
|[<span data-ttu-id="c9afc-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9afc-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9afc-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9afc-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c9afc-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c9afc-113">Members and methods</span></span>

| <span data-ttu-id="c9afc-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9afc-114">Member</span></span> | <span data-ttu-id="c9afc-115">種類</span><span class="sxs-lookup"><span data-stu-id="c9afc-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c9afc-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c9afc-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c9afc-117">Member</span><span class="sxs-lookup"><span data-stu-id="c9afc-117">Member</span></span> |
| [<span data-ttu-id="c9afc-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="c9afc-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="c9afc-119">Member</span><span class="sxs-lookup"><span data-stu-id="c9afc-119">Member</span></span> |
| [<span data-ttu-id="c9afc-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c9afc-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c9afc-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9afc-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c9afc-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="c9afc-122">Namespaces</span></span>

<span data-ttu-id="c9afc-123">[mailbox](office.context.mailbox.md): Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="c9afc-124">メンバー</span><span class="sxs-lookup"><span data-stu-id="c9afc-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="c9afc-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="c9afc-125">displayLanguage :String</span></span>

<span data-ttu-id="c9afc-126">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c9afc-127">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c9afc-128">型</span><span class="sxs-lookup"><span data-stu-id="c9afc-128">Type</span></span>

*   <span data-ttu-id="c9afc-129">String</span><span class="sxs-lookup"><span data-stu-id="c9afc-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9afc-130">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-130">Requirements</span></span>

|<span data-ttu-id="c9afc-131">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-131">Requirement</span></span>| <span data-ttu-id="c9afc-132">値</span><span class="sxs-lookup"><span data-stu-id="c9afc-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9afc-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9afc-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9afc-134">1.0</span><span class="sxs-lookup"><span data-stu-id="c9afc-134">1.0</span></span>|
|[<span data-ttu-id="c9afc-135">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9afc-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9afc-136">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9afc-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9afc-137">例</span><span class="sxs-lookup"><span data-stu-id="c9afc-137">Example</span></span>

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

---
---

####  <a name="officetheme-object"></a><span data-ttu-id="c9afc-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="c9afc-138">officeTheme :Object</span></span>

<span data-ttu-id="c9afc-139">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="c9afc-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c9afc-140">このメンバーは、Outlook for iOS または Outlook for Android ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c9afc-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c9afc-p102">Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。</span><span class="sxs-lookup"><span data-stu-id="c9afc-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c9afc-143">種類</span><span class="sxs-lookup"><span data-stu-id="c9afc-143">Type</span></span>

*   <span data-ttu-id="c9afc-144">Object</span><span class="sxs-lookup"><span data-stu-id="c9afc-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="c9afc-145">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c9afc-145">Properties:</span></span>

|<span data-ttu-id="c9afc-146">名前</span><span class="sxs-lookup"><span data-stu-id="c9afc-146">Name</span></span>| <span data-ttu-id="c9afc-147">種類</span><span class="sxs-lookup"><span data-stu-id="c9afc-147">Type</span></span>| <span data-ttu-id="c9afc-148">説明</span><span class="sxs-lookup"><span data-stu-id="c9afc-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c9afc-149">String</span><span class="sxs-lookup"><span data-stu-id="c9afc-149">String</span></span>|<span data-ttu-id="c9afc-150">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c9afc-151">String</span><span class="sxs-lookup"><span data-stu-id="c9afc-151">String</span></span>|<span data-ttu-id="c9afc-152">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c9afc-153">String</span><span class="sxs-lookup"><span data-stu-id="c9afc-153">String</span></span>|<span data-ttu-id="c9afc-154">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c9afc-155">String</span><span class="sxs-lookup"><span data-stu-id="c9afc-155">String</span></span>|<span data-ttu-id="c9afc-156">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c9afc-157">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-157">Requirements</span></span>

|<span data-ttu-id="c9afc-158">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-158">Requirement</span></span>| <span data-ttu-id="c9afc-159">値</span><span class="sxs-lookup"><span data-stu-id="c9afc-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9afc-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9afc-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9afc-161">1.3</span><span class="sxs-lookup"><span data-stu-id="c9afc-161">1.3</span></span>|
|[<span data-ttu-id="c9afc-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9afc-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9afc-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9afc-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9afc-164">例</span><span class="sxs-lookup"><span data-stu-id="c9afc-164">Example</span></span>

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

---
---

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="c9afc-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="c9afc-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="c9afc-166">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c9afc-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c9afc-167">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="c9afc-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c9afc-168">種類</span><span class="sxs-lookup"><span data-stu-id="c9afc-168">Type</span></span>

*   [<span data-ttu-id="c9afc-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c9afc-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c9afc-170">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-170">Requirements</span></span>

|<span data-ttu-id="c9afc-171">要件</span><span class="sxs-lookup"><span data-stu-id="c9afc-171">Requirement</span></span>| <span data-ttu-id="c9afc-172">値</span><span class="sxs-lookup"><span data-stu-id="c9afc-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9afc-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c9afc-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c9afc-174">1.0</span><span class="sxs-lookup"><span data-stu-id="c9afc-174">1.0</span></span>|
|[<span data-ttu-id="c9afc-175">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c9afc-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c9afc-176">制限あり</span><span class="sxs-lookup"><span data-stu-id="c9afc-176">Restricted</span></span>|
|[<span data-ttu-id="c9afc-177">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c9afc-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c9afc-178">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c9afc-178">Compose or Read</span></span>|
