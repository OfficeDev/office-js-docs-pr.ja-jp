---
title: Office コンテキスト要件セット1.4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ad1887f32568f30cb87e52dd1f9457be2022beb2
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127339"
---
# <a name="context"></a><span data-ttu-id="8aff8-102">context</span><span class="sxs-lookup"><span data-stu-id="8aff8-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="8aff8-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="8aff8-103">[Office](Office.md).context</span></span>

<span data-ttu-id="8aff8-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="8aff8-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8aff8-106">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-106">Requirements</span></span>

|<span data-ttu-id="8aff8-107">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-107">Requirement</span></span>| <span data-ttu-id="8aff8-108">値</span><span class="sxs-lookup"><span data-stu-id="8aff8-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aff8-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8aff8-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8aff8-110">1.0</span><span class="sxs-lookup"><span data-stu-id="8aff8-110">1.0</span></span>|
|[<span data-ttu-id="8aff8-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8aff8-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aff8-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8aff8-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="8aff8-113">名前空間</span><span class="sxs-lookup"><span data-stu-id="8aff8-113">Namespaces</span></span>

<span data-ttu-id="8aff8-114">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="8aff8-115">Members</span><span class="sxs-lookup"><span data-stu-id="8aff8-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="8aff8-116">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="8aff8-116">displayLanguage: String</span></span>

<span data-ttu-id="8aff8-117">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="8aff8-118">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="8aff8-119">型</span><span class="sxs-lookup"><span data-stu-id="8aff8-119">Type</span></span>

*   <span data-ttu-id="8aff8-120">String</span><span class="sxs-lookup"><span data-stu-id="8aff8-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8aff8-121">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-121">Requirements</span></span>

|<span data-ttu-id="8aff8-122">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-122">Requirement</span></span>| <span data-ttu-id="8aff8-123">値</span><span class="sxs-lookup"><span data-stu-id="8aff8-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aff8-124">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8aff8-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8aff8-125">1.0</span><span class="sxs-lookup"><span data-stu-id="8aff8-125">1.0</span></span>|
|[<span data-ttu-id="8aff8-126">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8aff8-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aff8-127">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8aff8-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aff8-128">例</span><span class="sxs-lookup"><span data-stu-id="8aff8-128">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="8aff8-129">officeTheme: オブジェクト</span><span class="sxs-lookup"><span data-stu-id="8aff8-129">officeTheme: Object</span></span>

<span data-ttu-id="8aff8-130">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="8aff8-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="8aff8-131">このメンバーは、iOS または Android の Outlook ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="8aff8-131">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8aff8-p102">Office テーマの色を使うと、**[ファイル] > [Office アカウント] > [Office テーマ UI]** によってユーザーが選択した現在の Office テーマに合わせてアドインの配色を調整できます。このテーマは Office ホスト アプリケーション全体に適用されます。Office テーマの色を使うことは、メール アドインと作業ウィンドウ アドインに適しています。</span><span class="sxs-lookup"><span data-stu-id="8aff8-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8aff8-134">型</span><span class="sxs-lookup"><span data-stu-id="8aff8-134">Type</span></span>

*   <span data-ttu-id="8aff8-135">Object</span><span class="sxs-lookup"><span data-stu-id="8aff8-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="8aff8-136">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="8aff8-136">Properties:</span></span>

|<span data-ttu-id="8aff8-137">名前</span><span class="sxs-lookup"><span data-stu-id="8aff8-137">Name</span></span>| <span data-ttu-id="8aff8-138">種類</span><span class="sxs-lookup"><span data-stu-id="8aff8-138">Type</span></span>| <span data-ttu-id="8aff8-139">説明</span><span class="sxs-lookup"><span data-stu-id="8aff8-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="8aff8-140">String</span><span class="sxs-lookup"><span data-stu-id="8aff8-140">String</span></span>|<span data-ttu-id="8aff8-141">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="8aff8-142">String</span><span class="sxs-lookup"><span data-stu-id="8aff8-142">String</span></span>|<span data-ttu-id="8aff8-143">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="8aff8-144">String</span><span class="sxs-lookup"><span data-stu-id="8aff8-144">String</span></span>|<span data-ttu-id="8aff8-145">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="8aff8-146">String</span><span class="sxs-lookup"><span data-stu-id="8aff8-146">String</span></span>|<span data-ttu-id="8aff8-147">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8aff8-148">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-148">Requirements</span></span>

|<span data-ttu-id="8aff8-149">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-149">Requirement</span></span>| <span data-ttu-id="8aff8-150">値</span><span class="sxs-lookup"><span data-stu-id="8aff8-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aff8-151">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8aff8-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8aff8-152">1.3</span><span class="sxs-lookup"><span data-stu-id="8aff8-152">1.3</span></span>|
|[<span data-ttu-id="8aff8-153">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8aff8-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aff8-154">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8aff8-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8aff8-155">例</span><span class="sxs-lookup"><span data-stu-id="8aff8-155">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="8aff8-156">roamingSettings: [roamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="8aff8-156">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="8aff8-157">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="8aff8-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8aff8-158">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="8aff8-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8aff8-159">型</span><span class="sxs-lookup"><span data-stu-id="8aff8-159">Type</span></span>

*   [<span data-ttu-id="8aff8-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8aff8-160">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8aff8-161">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-161">Requirements</span></span>

|<span data-ttu-id="8aff8-162">要件</span><span class="sxs-lookup"><span data-stu-id="8aff8-162">Requirement</span></span>| <span data-ttu-id="8aff8-163">値</span><span class="sxs-lookup"><span data-stu-id="8aff8-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="8aff8-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="8aff8-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8aff8-165">1.0</span><span class="sxs-lookup"><span data-stu-id="8aff8-165">1.0</span></span>|
|[<span data-ttu-id="8aff8-166">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="8aff8-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8aff8-167">制限あり</span><span class="sxs-lookup"><span data-stu-id="8aff8-167">Restricted</span></span>|
|[<span data-ttu-id="8aff8-168">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="8aff8-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8aff8-169">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="8aff8-169">Compose or Read</span></span>|
