---
title: Office コンテキスト-プレビュー要件セット
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629189"
---
# <a name="context"></a><span data-ttu-id="57ef0-102">context</span><span class="sxs-lookup"><span data-stu-id="57ef0-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="57ef0-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="57ef0-103">[Office](Office.md).context</span></span>

<span data-ttu-id="57ef0-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="57ef0-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="57ef0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-106">Requirements</span></span>

|<span data-ttu-id="57ef0-107">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-107">Requirement</span></span>| <span data-ttu-id="57ef0-108">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-110">1.0</span></span>|
|[<span data-ttu-id="57ef0-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="57ef0-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57ef0-113">Properties</span></span>

| <span data-ttu-id="57ef0-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="57ef0-114">Property</span></span> | <span data-ttu-id="57ef0-115">モード</span><span class="sxs-lookup"><span data-stu-id="57ef0-115">Modes</span></span> | <span data-ttu-id="57ef0-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="57ef0-116">Return type</span></span> | <span data-ttu-id="57ef0-117">最小値</span><span class="sxs-lookup"><span data-stu-id="57ef0-117">Minimum</span></span><br><span data-ttu-id="57ef0-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="57ef0-118">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="57ef0-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="57ef0-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="57ef0-120">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-120">Compose</span></span><br><span data-ttu-id="57ef0-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-121">Read</span></span> | <span data-ttu-id="57ef0-122">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-122">String</span></span> | <span data-ttu-id="57ef0-123">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-123">1.0</span></span> |
| [<span data-ttu-id="57ef0-124">ダン</span><span class="sxs-lookup"><span data-stu-id="57ef0-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="57ef0-125">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-125">Compose</span></span><br><span data-ttu-id="57ef0-126">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-126">Read</span></span> | [<span data-ttu-id="57ef0-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="57ef0-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation) | <span data-ttu-id="57ef0-128">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-128">1.0</span></span> |
| [<span data-ttu-id="57ef0-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="57ef0-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="57ef0-130">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-130">Compose</span></span><br><span data-ttu-id="57ef0-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-131">Read</span></span> | <span data-ttu-id="57ef0-132">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-132">String</span></span> | <span data-ttu-id="57ef0-133">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-133">1.0</span></span> |
| [<span data-ttu-id="57ef0-134">主催</span><span class="sxs-lookup"><span data-stu-id="57ef0-134">host</span></span>](#host-hosttype) | <span data-ttu-id="57ef0-135">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-135">Compose</span></span><br><span data-ttu-id="57ef0-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-136">Read</span></span> | [<span data-ttu-id="57ef0-137">HostType</span><span class="sxs-lookup"><span data-stu-id="57ef0-137">HostType</span></span>](/javascript/api/office/office.hosttype) | <span data-ttu-id="57ef0-138">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-138">1.0</span></span> |
| [<span data-ttu-id="57ef0-139">officeTheme</span><span class="sxs-lookup"><span data-stu-id="57ef0-139">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="57ef0-140">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-140">Compose</span></span><br><span data-ttu-id="57ef0-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-141">Read</span></span> | [<span data-ttu-id="57ef0-142">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="57ef0-142">OfficeTheme</span></span>](/javascript/api/office/office.officetheme) | <span data-ttu-id="57ef0-143">プレビュー</span><span class="sxs-lookup"><span data-stu-id="57ef0-143">Preview</span></span> |
| [<span data-ttu-id="57ef0-144">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="57ef0-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="57ef0-145">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-145">Compose</span></span><br><span data-ttu-id="57ef0-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-146">Read</span></span> | [<span data-ttu-id="57ef0-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="57ef0-147">PlatformType</span></span>](/javascript/api/office/office.platformtype) | <span data-ttu-id="57ef0-148">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-148">1.0</span></span> |
| [<span data-ttu-id="57ef0-149">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="57ef0-150">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-150">Compose</span></span><br><span data-ttu-id="57ef0-151">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-151">Read</span></span> | [<span data-ttu-id="57ef0-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="57ef0-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport) | <span data-ttu-id="57ef0-153">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-153">1.0</span></span> |
| [<span data-ttu-id="57ef0-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="57ef0-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="57ef0-155">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-155">Compose</span></span><br><span data-ttu-id="57ef0-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-156">Read</span></span> | [<span data-ttu-id="57ef0-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="57ef0-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings) | <span data-ttu-id="57ef0-158">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-158">1.0</span></span> |
| [<span data-ttu-id="57ef0-159">UI</span><span class="sxs-lookup"><span data-stu-id="57ef0-159">ui</span></span>](#ui-ui) | <span data-ttu-id="57ef0-160">作成</span><span class="sxs-lookup"><span data-stu-id="57ef0-160">Compose</span></span><br><span data-ttu-id="57ef0-161">読み取り</span><span class="sxs-lookup"><span data-stu-id="57ef0-161">Read</span></span> | [<span data-ttu-id="57ef0-162">UI</span><span class="sxs-lookup"><span data-stu-id="57ef0-162">UI</span></span>](/javascript/api/office/office.ui) | <span data-ttu-id="57ef0-163">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-163">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="57ef0-164">名前空間</span><span class="sxs-lookup"><span data-stu-id="57ef0-164">Namespaces</span></span>

<span data-ttu-id="57ef0-165">[auth](/javascript/api/office/office.auth):[シングルサインオン (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token)のサポートを提供します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-165">[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span></span>

<span data-ttu-id="57ef0-166">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-166">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

## <a name="property-details"></a><span data-ttu-id="57ef0-167">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="57ef0-167">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="57ef0-168">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="57ef0-168">contentLanguage: String</span></span>

<span data-ttu-id="57ef0-169">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-169">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="57ef0-170">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-170">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-171">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-171">Type</span></span>

*   <span data-ttu-id="57ef0-172">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-172">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57ef0-173">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-173">Requirements</span></span>

|<span data-ttu-id="57ef0-174">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-174">Requirement</span></span>| <span data-ttu-id="57ef0-175">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-177">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-177">1.0</span></span>|
|[<span data-ttu-id="57ef0-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-180">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-180">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
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

<br>

---
---

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="57ef0-181">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="57ef0-181">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="57ef0-182">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-182">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-183">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-183">Type</span></span>

*   [<span data-ttu-id="57ef0-184">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="57ef0-184">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="57ef0-185">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-185">Requirements</span></span>

|<span data-ttu-id="57ef0-186">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-186">Requirement</span></span>| <span data-ttu-id="57ef0-187">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-188">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-189">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-189">1.0</span></span>|
|[<span data-ttu-id="57ef0-190">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-190">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-191">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-191">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-192">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-192">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="57ef0-193">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="57ef0-193">displayLanguage: String</span></span>

<span data-ttu-id="57ef0-194">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-194">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="57ef0-195">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-196">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-196">Type</span></span>

*   <span data-ttu-id="57ef0-197">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-197">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="57ef0-198">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-198">Requirements</span></span>

|<span data-ttu-id="57ef0-199">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-199">Requirement</span></span>| <span data-ttu-id="57ef0-200">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-201">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-202">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-202">1.0</span></span>|
|[<span data-ttu-id="57ef0-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-203">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-205">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-205">Example</span></span>

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

<br>

---
---

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="57ef0-206">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="57ef0-206">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="57ef0-207">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-207">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-208">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-208">Type</span></span>

*   [<span data-ttu-id="57ef0-209">HostType</span><span class="sxs-lookup"><span data-stu-id="57ef0-209">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="57ef0-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-210">Requirements</span></span>

|<span data-ttu-id="57ef0-211">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-211">Requirement</span></span>| <span data-ttu-id="57ef0-212">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-213">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-214">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-214">1.0</span></span>|
|[<span data-ttu-id="57ef0-215">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-216">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-216">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-217">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-217">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="57ef0-218">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="57ef0-218">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="57ef0-219">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="57ef0-219">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="57ef0-220">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="57ef0-220">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="57ef0-221">Office テーマの色を使用すると、アドインの配色を、[**ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office ホストアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="57ef0-221">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="57ef0-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="57ef0-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-223">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-223">Type</span></span>

*   [<span data-ttu-id="57ef0-224">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="57ef0-224">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="57ef0-225">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="57ef0-225">Properties:</span></span>

|<span data-ttu-id="57ef0-226">名前</span><span class="sxs-lookup"><span data-stu-id="57ef0-226">Name</span></span>| <span data-ttu-id="57ef0-227">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-227">Type</span></span>| <span data-ttu-id="57ef0-228">説明</span><span class="sxs-lookup"><span data-stu-id="57ef0-228">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="57ef0-229">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-229">String</span></span>|<span data-ttu-id="57ef0-230">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-230">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="57ef0-231">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-231">String</span></span>|<span data-ttu-id="57ef0-232">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-232">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="57ef0-233">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-233">String</span></span>|<span data-ttu-id="57ef0-234">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-234">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="57ef0-235">String</span><span class="sxs-lookup"><span data-stu-id="57ef0-235">String</span></span>|<span data-ttu-id="57ef0-236">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-236">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57ef0-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-237">Requirements</span></span>

|<span data-ttu-id="57ef0-238">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-238">Requirement</span></span>| <span data-ttu-id="57ef0-239">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-240">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-241">プレビュー</span><span class="sxs-lookup"><span data-stu-id="57ef0-241">Preview</span></span>|
|[<span data-ttu-id="57ef0-242">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-243">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-243">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-244">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-244">Example</span></span>

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

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="57ef0-245">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="57ef0-245">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="57ef0-246">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-246">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-247">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-247">Type</span></span>

*   [<span data-ttu-id="57ef0-248">PlatformType</span><span class="sxs-lookup"><span data-stu-id="57ef0-248">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="57ef0-249">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-249">Requirements</span></span>

|<span data-ttu-id="57ef0-250">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-250">Requirement</span></span>| <span data-ttu-id="57ef0-251">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-252">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-252">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-253">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-253">1.0</span></span>|
|[<span data-ttu-id="57ef0-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-255">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-256">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-256">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="57ef0-257">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="57ef0-257">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="57ef0-258">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-258">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-259">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-259">Type</span></span>

*   [<span data-ttu-id="57ef0-260">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="57ef0-260">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="57ef0-261">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-261">Requirements</span></span>

|<span data-ttu-id="57ef0-262">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-262">Requirement</span></span>| <span data-ttu-id="57ef0-263">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-264">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-265">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-265">1.0</span></span>|
|[<span data-ttu-id="57ef0-266">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-266">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-267">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-267">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="57ef0-268">例</span><span class="sxs-lookup"><span data-stu-id="57ef0-268">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="57ef0-269">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="57ef0-269">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="57ef0-270">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-270">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="57ef0-271">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="57ef0-271">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-272">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-272">Type</span></span>

*   [<span data-ttu-id="57ef0-273">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="57ef0-273">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="57ef0-274">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-274">Requirements</span></span>

|<span data-ttu-id="57ef0-275">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-275">Requirement</span></span>| <span data-ttu-id="57ef0-276">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-277">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-278">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-278">1.0</span></span>|
|[<span data-ttu-id="57ef0-279">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="57ef0-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="57ef0-280">制限あり</span><span class="sxs-lookup"><span data-stu-id="57ef0-280">Restricted</span></span>|
|[<span data-ttu-id="57ef0-281">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-282">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-282">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="57ef0-283">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="57ef0-283">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="57ef0-284">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="57ef0-284">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="57ef0-285">型</span><span class="sxs-lookup"><span data-stu-id="57ef0-285">Type</span></span>

*   [<span data-ttu-id="57ef0-286">UI</span><span class="sxs-lookup"><span data-stu-id="57ef0-286">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="57ef0-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="57ef0-287">Requirements</span></span>

|<span data-ttu-id="57ef0-288">要件</span><span class="sxs-lookup"><span data-stu-id="57ef0-288">Requirement</span></span>| <span data-ttu-id="57ef0-289">値</span><span class="sxs-lookup"><span data-stu-id="57ef0-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ef0-290">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="57ef0-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="57ef0-291">1.0</span><span class="sxs-lookup"><span data-stu-id="57ef0-291">1.0</span></span>|
|[<span data-ttu-id="57ef0-292">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="57ef0-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ef0-293">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="57ef0-293">Compose or Read</span></span>|
