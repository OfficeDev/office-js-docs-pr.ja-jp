---
title: Office コンテキスト要件セット1.7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ff816b3bb51ebb5dc8ef124af8488405fdc3fd39
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127136"
---
# <a name="context"></a><span data-ttu-id="c5d9d-102">context</span><span class="sxs-lookup"><span data-stu-id="c5d9d-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="c5d9d-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="c5d9d-103">[Office](Office.md).context</span></span>

<span data-ttu-id="c5d9d-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5d9d-106">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-106">Requirements</span></span>

|<span data-ttu-id="c5d9d-107">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-107">Requirement</span></span>| <span data-ttu-id="c5d9d-108">値</span><span class="sxs-lookup"><span data-stu-id="c5d9d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5d9d-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5d9d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5d9d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c5d9d-110">1.0</span></span>|
|[<span data-ttu-id="c5d9d-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5d9d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5d9d-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5d9d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c5d9d-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="c5d9d-113">Members and methods</span></span>

| <span data-ttu-id="c5d9d-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="c5d9d-114">Member</span></span> | <span data-ttu-id="c5d9d-115">種類</span><span class="sxs-lookup"><span data-stu-id="c5d9d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c5d9d-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c5d9d-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c5d9d-117">Member</span><span class="sxs-lookup"><span data-stu-id="c5d9d-117">Member</span></span> |
| [<span data-ttu-id="c5d9d-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="c5d9d-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="c5d9d-119">Member</span><span class="sxs-lookup"><span data-stu-id="c5d9d-119">Member</span></span> |
| [<span data-ttu-id="c5d9d-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5d9d-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c5d9d-121">メンバー</span><span class="sxs-lookup"><span data-stu-id="c5d9d-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c5d9d-122">名前空間</span><span class="sxs-lookup"><span data-stu-id="c5d9d-122">Namespaces</span></span>

<span data-ttu-id="c5d9d-123">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="c5d9d-124">Members</span><span class="sxs-lookup"><span data-stu-id="c5d9d-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="c5d9d-125">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-125">displayLanguage: String</span></span>

<span data-ttu-id="c5d9d-126">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c5d9d-127">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5d9d-128">型</span><span class="sxs-lookup"><span data-stu-id="c5d9d-128">Type</span></span>

*   <span data-ttu-id="c5d9d-129">String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5d9d-130">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-130">Requirements</span></span>

|<span data-ttu-id="c5d9d-131">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-131">Requirement</span></span>| <span data-ttu-id="c5d9d-132">値</span><span class="sxs-lookup"><span data-stu-id="c5d9d-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5d9d-133">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5d9d-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5d9d-134">1.0</span><span class="sxs-lookup"><span data-stu-id="c5d9d-134">1.0</span></span>|
|[<span data-ttu-id="c5d9d-135">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5d9d-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5d9d-136">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5d9d-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5d9d-137">例</span><span class="sxs-lookup"><span data-stu-id="c5d9d-137">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="c5d9d-138">officeTheme: オブジェクト</span><span class="sxs-lookup"><span data-stu-id="c5d9d-138">officeTheme: Object</span></span>

<span data-ttu-id="c5d9d-139">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c5d9d-140">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-140">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="c5d9d-p102">Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c5d9d-143">型</span><span class="sxs-lookup"><span data-stu-id="c5d9d-143">Type</span></span>

*   <span data-ttu-id="c5d9d-144">Object</span><span class="sxs-lookup"><span data-stu-id="c5d9d-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="c5d9d-145">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c5d9d-145">Properties:</span></span>

|<span data-ttu-id="c5d9d-146">名前</span><span class="sxs-lookup"><span data-stu-id="c5d9d-146">Name</span></span>| <span data-ttu-id="c5d9d-147">種類</span><span class="sxs-lookup"><span data-stu-id="c5d9d-147">Type</span></span>| <span data-ttu-id="c5d9d-148">説明</span><span class="sxs-lookup"><span data-stu-id="c5d9d-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c5d9d-149">String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-149">String</span></span>|<span data-ttu-id="c5d9d-150">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c5d9d-151">String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-151">String</span></span>|<span data-ttu-id="c5d9d-152">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c5d9d-153">String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-153">String</span></span>|<span data-ttu-id="c5d9d-154">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c5d9d-155">String</span><span class="sxs-lookup"><span data-stu-id="c5d9d-155">String</span></span>|<span data-ttu-id="c5d9d-156">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c5d9d-157">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-157">Requirements</span></span>

|<span data-ttu-id="c5d9d-158">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-158">Requirement</span></span>| <span data-ttu-id="c5d9d-159">値</span><span class="sxs-lookup"><span data-stu-id="c5d9d-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5d9d-160">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5d9d-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5d9d-161">1.3</span><span class="sxs-lookup"><span data-stu-id="c5d9d-161">1.3</span></span>|
|[<span data-ttu-id="c5d9d-162">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5d9d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5d9d-163">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5d9d-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5d9d-164">例</span><span class="sxs-lookup"><span data-stu-id="c5d9d-164">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="c5d9d-165">roamingSettings: [roamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="c5d9d-165">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="c5d9d-166">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c5d9d-167">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="c5d9d-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c5d9d-168">型</span><span class="sxs-lookup"><span data-stu-id="c5d9d-168">Type</span></span>

*   [<span data-ttu-id="c5d9d-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5d9d-169">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c5d9d-170">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-170">Requirements</span></span>

|<span data-ttu-id="c5d9d-171">要件</span><span class="sxs-lookup"><span data-stu-id="c5d9d-171">Requirement</span></span>| <span data-ttu-id="c5d9d-172">値</span><span class="sxs-lookup"><span data-stu-id="c5d9d-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5d9d-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5d9d-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5d9d-174">1.0</span><span class="sxs-lookup"><span data-stu-id="c5d9d-174">1.0</span></span>|
|[<span data-ttu-id="c5d9d-175">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c5d9d-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c5d9d-176">制限あり</span><span class="sxs-lookup"><span data-stu-id="c5d9d-176">Restricted</span></span>|
|[<span data-ttu-id="c5d9d-177">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5d9d-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5d9d-178">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5d9d-178">Compose or Read</span></span>|
