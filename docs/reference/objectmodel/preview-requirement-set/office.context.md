---
title: Office コンテキスト-プレビュー要件セット
description: メールボックス API プレビュー要件セットを使用して Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: c61769cb1ae98097ffabb8b3ef19b2f82257c2b1
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890866"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="56017-103">コンテキスト (メールボックスプレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="56017-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="56017-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="56017-104">[Office](office.md).context</span></span>

<span data-ttu-id="56017-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="56017-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="56017-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="56017-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="56017-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-107">Requirements</span></span>

|<span data-ttu-id="56017-108">要件</span><span class="sxs-lookup"><span data-stu-id="56017-108">Requirement</span></span>| <span data-ttu-id="56017-109">値</span><span class="sxs-lookup"><span data-stu-id="56017-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-111">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-111">1.1</span></span>|
|[<span data-ttu-id="56017-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="56017-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="56017-114">Properties</span></span>

| <span data-ttu-id="56017-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="56017-115">Property</span></span> | <span data-ttu-id="56017-116">モード</span><span class="sxs-lookup"><span data-stu-id="56017-116">Modes</span></span> | <span data-ttu-id="56017-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="56017-117">Return type</span></span> | <span data-ttu-id="56017-118">最小値</span><span class="sxs-lookup"><span data-stu-id="56017-118">Minimum</span></span><br><span data-ttu-id="56017-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="56017-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="56017-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="56017-120">auth</span></span>](#auth-auth) | <span data-ttu-id="56017-121">作成</span><span class="sxs-lookup"><span data-stu-id="56017-121">Compose</span></span><br><span data-ttu-id="56017-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-122">Read</span></span> | [<span data-ttu-id="56017-123">Auth</span><span class="sxs-lookup"><span data-stu-id="56017-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="56017-124">プレビュー</span><span class="sxs-lookup"><span data-stu-id="56017-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="56017-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="56017-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="56017-126">作成</span><span class="sxs-lookup"><span data-stu-id="56017-126">Compose</span></span><br><span data-ttu-id="56017-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-127">Read</span></span> | <span data-ttu-id="56017-128">String</span><span class="sxs-lookup"><span data-stu-id="56017-128">String</span></span> | [<span data-ttu-id="56017-129">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-130">ダン</span><span class="sxs-lookup"><span data-stu-id="56017-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="56017-131">作成</span><span class="sxs-lookup"><span data-stu-id="56017-131">Compose</span></span><br><span data-ttu-id="56017-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-132">Read</span></span> | [<span data-ttu-id="56017-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="56017-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="56017-134">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="56017-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="56017-136">作成</span><span class="sxs-lookup"><span data-stu-id="56017-136">Compose</span></span><br><span data-ttu-id="56017-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-137">Read</span></span> | <span data-ttu-id="56017-138">String</span><span class="sxs-lookup"><span data-stu-id="56017-138">String</span></span> | [<span data-ttu-id="56017-139">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-140">主催</span><span class="sxs-lookup"><span data-stu-id="56017-140">host</span></span>](#host-hosttype) | <span data-ttu-id="56017-141">作成</span><span class="sxs-lookup"><span data-stu-id="56017-141">Compose</span></span><br><span data-ttu-id="56017-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-142">Read</span></span> | [<span data-ttu-id="56017-143">HostType</span><span class="sxs-lookup"><span data-stu-id="56017-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="56017-144">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="56017-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="56017-146">作成</span><span class="sxs-lookup"><span data-stu-id="56017-146">Compose</span></span><br><span data-ttu-id="56017-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-147">Read</span></span> | [<span data-ttu-id="56017-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="56017-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="56017-149">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="56017-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="56017-151">作成</span><span class="sxs-lookup"><span data-stu-id="56017-151">Compose</span></span><br><span data-ttu-id="56017-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-152">Read</span></span> | [<span data-ttu-id="56017-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="56017-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="56017-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="56017-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="56017-155">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="56017-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="56017-156">作成</span><span class="sxs-lookup"><span data-stu-id="56017-156">Compose</span></span><br><span data-ttu-id="56017-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-157">Read</span></span> | [<span data-ttu-id="56017-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="56017-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="56017-159">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-160">要件</span><span class="sxs-lookup"><span data-stu-id="56017-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="56017-161">作成</span><span class="sxs-lookup"><span data-stu-id="56017-161">Compose</span></span><br><span data-ttu-id="56017-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-162">Read</span></span> | [<span data-ttu-id="56017-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="56017-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="56017-164">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="56017-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="56017-166">作成</span><span class="sxs-lookup"><span data-stu-id="56017-166">Compose</span></span><br><span data-ttu-id="56017-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-167">Read</span></span> | [<span data-ttu-id="56017-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="56017-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="56017-169">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="56017-170">UI</span><span class="sxs-lookup"><span data-stu-id="56017-170">ui</span></span>](#ui-ui) | <span data-ttu-id="56017-171">作成</span><span class="sxs-lookup"><span data-stu-id="56017-171">Compose</span></span><br><span data-ttu-id="56017-172">読み取り</span><span class="sxs-lookup"><span data-stu-id="56017-172">Read</span></span> | [<span data-ttu-id="56017-173">UI</span><span class="sxs-lookup"><span data-stu-id="56017-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="56017-174">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="56017-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="56017-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="56017-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="56017-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="56017-177">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office ホストがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="56017-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="56017-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="56017-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-179">型</span><span class="sxs-lookup"><span data-stu-id="56017-179">Type</span></span>

*   [<span data-ttu-id="56017-180">Auth</span><span class="sxs-lookup"><span data-stu-id="56017-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="56017-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-181">Requirements</span></span>

|<span data-ttu-id="56017-182">要件</span><span class="sxs-lookup"><span data-stu-id="56017-182">Requirement</span></span>| <span data-ttu-id="56017-183">値</span><span class="sxs-lookup"><span data-stu-id="56017-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="56017-185">Preview</span></span>|
|[<span data-ttu-id="56017-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-188">例</span><span class="sxs-lookup"><span data-stu-id="56017-188">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="56017-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="56017-189">contentLanguage: String</span></span>

<span data-ttu-id="56017-190">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="56017-191">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="56017-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-192">型</span><span class="sxs-lookup"><span data-stu-id="56017-192">Type</span></span>

*   <span data-ttu-id="56017-193">String</span><span class="sxs-lookup"><span data-stu-id="56017-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="56017-194">要件</span><span class="sxs-lookup"><span data-stu-id="56017-194">Requirements</span></span>

|<span data-ttu-id="56017-195">要件</span><span class="sxs-lookup"><span data-stu-id="56017-195">Requirement</span></span>| <span data-ttu-id="56017-196">値</span><span class="sxs-lookup"><span data-stu-id="56017-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-198">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-198">1.1</span></span>|
|[<span data-ttu-id="56017-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-201">例</span><span class="sxs-lookup"><span data-stu-id="56017-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="56017-202">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="56017-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="56017-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-204">型</span><span class="sxs-lookup"><span data-stu-id="56017-204">Type</span></span>

*   [<span data-ttu-id="56017-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="56017-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="56017-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-206">Requirements</span></span>

|<span data-ttu-id="56017-207">要件</span><span class="sxs-lookup"><span data-stu-id="56017-207">Requirement</span></span>| <span data-ttu-id="56017-208">値</span><span class="sxs-lookup"><span data-stu-id="56017-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-210">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-210">1.1</span></span>|
|[<span data-ttu-id="56017-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-213">例</span><span class="sxs-lookup"><span data-stu-id="56017-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="56017-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="56017-214">displayLanguage: String</span></span>

<span data-ttu-id="56017-215">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="56017-216">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="56017-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-217">型</span><span class="sxs-lookup"><span data-stu-id="56017-217">Type</span></span>

*   <span data-ttu-id="56017-218">String</span><span class="sxs-lookup"><span data-stu-id="56017-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="56017-219">要件</span><span class="sxs-lookup"><span data-stu-id="56017-219">Requirements</span></span>

|<span data-ttu-id="56017-220">要件</span><span class="sxs-lookup"><span data-stu-id="56017-220">Requirement</span></span>| <span data-ttu-id="56017-221">値</span><span class="sxs-lookup"><span data-stu-id="56017-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-223">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-223">1.1</span></span>|
|[<span data-ttu-id="56017-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-226">例</span><span class="sxs-lookup"><span data-stu-id="56017-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="56017-227">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="56017-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="56017-228">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-229">型</span><span class="sxs-lookup"><span data-stu-id="56017-229">Type</span></span>

*   [<span data-ttu-id="56017-230">HostType</span><span class="sxs-lookup"><span data-stu-id="56017-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="56017-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-231">Requirements</span></span>

|<span data-ttu-id="56017-232">要件</span><span class="sxs-lookup"><span data-stu-id="56017-232">Requirement</span></span>| <span data-ttu-id="56017-233">値</span><span class="sxs-lookup"><span data-stu-id="56017-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-235">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-235">1.1</span></span>|
|[<span data-ttu-id="56017-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-238">例</span><span class="sxs-lookup"><span data-stu-id="56017-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="56017-239">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="56017-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="56017-240">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="56017-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="56017-241">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="56017-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="56017-242">Office テーマの色を使用すると、アドインの配色を、[**ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office ホストアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="56017-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="56017-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="56017-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-244">型</span><span class="sxs-lookup"><span data-stu-id="56017-244">Type</span></span>

*   [<span data-ttu-id="56017-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="56017-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="56017-246">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="56017-246">Properties:</span></span>

|<span data-ttu-id="56017-247">名前</span><span class="sxs-lookup"><span data-stu-id="56017-247">Name</span></span>| <span data-ttu-id="56017-248">種類</span><span class="sxs-lookup"><span data-stu-id="56017-248">Type</span></span>| <span data-ttu-id="56017-249">説明</span><span class="sxs-lookup"><span data-stu-id="56017-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="56017-250">String</span><span class="sxs-lookup"><span data-stu-id="56017-250">String</span></span>|<span data-ttu-id="56017-251">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="56017-252">String</span><span class="sxs-lookup"><span data-stu-id="56017-252">String</span></span>|<span data-ttu-id="56017-253">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="56017-254">String</span><span class="sxs-lookup"><span data-stu-id="56017-254">String</span></span>|<span data-ttu-id="56017-255">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="56017-256">String</span><span class="sxs-lookup"><span data-stu-id="56017-256">String</span></span>|<span data-ttu-id="56017-257">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="56017-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-258">Requirements</span></span>

|<span data-ttu-id="56017-259">要件</span><span class="sxs-lookup"><span data-stu-id="56017-259">Requirement</span></span>| <span data-ttu-id="56017-260">値</span><span class="sxs-lookup"><span data-stu-id="56017-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-262">プレビュー</span><span class="sxs-lookup"><span data-stu-id="56017-262">Preview</span></span>|
|[<span data-ttu-id="56017-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-265">例</span><span class="sxs-lookup"><span data-stu-id="56017-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="56017-266">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="56017-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="56017-267">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="56017-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-268">型</span><span class="sxs-lookup"><span data-stu-id="56017-268">Type</span></span>

*   [<span data-ttu-id="56017-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="56017-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="56017-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-270">Requirements</span></span>

|<span data-ttu-id="56017-271">要件</span><span class="sxs-lookup"><span data-stu-id="56017-271">Requirement</span></span>| <span data-ttu-id="56017-272">値</span><span class="sxs-lookup"><span data-stu-id="56017-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-274">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-274">1.1</span></span>|
|[<span data-ttu-id="56017-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-277">例</span><span class="sxs-lookup"><span data-stu-id="56017-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="56017-278">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="56017-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="56017-279">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="56017-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-280">型</span><span class="sxs-lookup"><span data-stu-id="56017-280">Type</span></span>

*   [<span data-ttu-id="56017-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="56017-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="56017-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-282">Requirements</span></span>

|<span data-ttu-id="56017-283">要件</span><span class="sxs-lookup"><span data-stu-id="56017-283">Requirement</span></span>| <span data-ttu-id="56017-284">値</span><span class="sxs-lookup"><span data-stu-id="56017-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-286">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-286">1.1</span></span>|
|[<span data-ttu-id="56017-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-288">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="56017-289">例</span><span class="sxs-lookup"><span data-stu-id="56017-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="56017-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="56017-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="56017-291">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="56017-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="56017-292">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="56017-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-293">型</span><span class="sxs-lookup"><span data-stu-id="56017-293">Type</span></span>

*   [<span data-ttu-id="56017-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="56017-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="56017-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-295">Requirements</span></span>

|<span data-ttu-id="56017-296">要件</span><span class="sxs-lookup"><span data-stu-id="56017-296">Requirement</span></span>| <span data-ttu-id="56017-297">値</span><span class="sxs-lookup"><span data-stu-id="56017-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-299">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-299">1.1</span></span>|
|[<span data-ttu-id="56017-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="56017-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="56017-301">制限あり</span><span class="sxs-lookup"><span data-stu-id="56017-301">Restricted</span></span>|
|[<span data-ttu-id="56017-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-303">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="56017-304">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="56017-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="56017-305">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="56017-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="56017-306">型</span><span class="sxs-lookup"><span data-stu-id="56017-306">Type</span></span>

*   [<span data-ttu-id="56017-307">UI</span><span class="sxs-lookup"><span data-stu-id="56017-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="56017-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="56017-308">Requirements</span></span>

|<span data-ttu-id="56017-309">要件</span><span class="sxs-lookup"><span data-stu-id="56017-309">Requirement</span></span>| <span data-ttu-id="56017-310">値</span><span class="sxs-lookup"><span data-stu-id="56017-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="56017-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="56017-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="56017-312">1.1</span><span class="sxs-lookup"><span data-stu-id="56017-312">1.1</span></span>|
|[<span data-ttu-id="56017-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="56017-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="56017-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="56017-314">Compose or Read</span></span>|
