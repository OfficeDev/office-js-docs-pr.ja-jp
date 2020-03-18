---
title: Office コンテキスト-プレビュー要件セット
description: Outlook アドイン API (Mailbox API Preview バージョン) の Outlook コンテキストオブジェクトのオブジェクトモデル。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 409f0a5b46eba667f79228f45081c160c3c3ce7f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717804"
---
# <a name="context"></a><span data-ttu-id="05a29-103">context</span><span class="sxs-lookup"><span data-stu-id="05a29-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="05a29-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="05a29-104">[Office](office.md).context</span></span>

<span data-ttu-id="05a29-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="05a29-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="05a29-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="05a29-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="05a29-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-107">Requirements</span></span>

|<span data-ttu-id="05a29-108">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-108">Requirement</span></span>| <span data-ttu-id="05a29-109">値</span><span class="sxs-lookup"><span data-stu-id="05a29-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-111">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-111">1.1</span></span>|
|[<span data-ttu-id="05a29-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="05a29-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="05a29-114">Properties</span></span>

| <span data-ttu-id="05a29-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="05a29-115">Property</span></span> | <span data-ttu-id="05a29-116">モード</span><span class="sxs-lookup"><span data-stu-id="05a29-116">Modes</span></span> | <span data-ttu-id="05a29-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="05a29-117">Return type</span></span> | <span data-ttu-id="05a29-118">最小値</span><span class="sxs-lookup"><span data-stu-id="05a29-118">Minimum</span></span><br><span data-ttu-id="05a29-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="05a29-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="05a29-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="05a29-120">auth</span></span>](#auth-auth) | <span data-ttu-id="05a29-121">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-121">Compose</span></span><br><span data-ttu-id="05a29-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-122">Read</span></span> | [<span data-ttu-id="05a29-123">Auth</span><span class="sxs-lookup"><span data-stu-id="05a29-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="05a29-124">プレビュー</span><span class="sxs-lookup"><span data-stu-id="05a29-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="05a29-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="05a29-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="05a29-126">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-126">Compose</span></span><br><span data-ttu-id="05a29-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-127">Read</span></span> | <span data-ttu-id="05a29-128">文字列</span><span class="sxs-lookup"><span data-stu-id="05a29-128">String</span></span> | [<span data-ttu-id="05a29-129">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-130">ダン</span><span class="sxs-lookup"><span data-stu-id="05a29-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="05a29-131">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-131">Compose</span></span><br><span data-ttu-id="05a29-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-132">Read</span></span> | [<span data-ttu-id="05a29-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="05a29-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="05a29-134">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="05a29-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="05a29-136">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-136">Compose</span></span><br><span data-ttu-id="05a29-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-137">Read</span></span> | <span data-ttu-id="05a29-138">文字列</span><span class="sxs-lookup"><span data-stu-id="05a29-138">String</span></span> | [<span data-ttu-id="05a29-139">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-140">主催</span><span class="sxs-lookup"><span data-stu-id="05a29-140">host</span></span>](#host-hosttype) | <span data-ttu-id="05a29-141">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-141">Compose</span></span><br><span data-ttu-id="05a29-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-142">Read</span></span> | [<span data-ttu-id="05a29-143">HostType</span><span class="sxs-lookup"><span data-stu-id="05a29-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="05a29-144">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="05a29-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="05a29-146">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-146">Compose</span></span><br><span data-ttu-id="05a29-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-147">Read</span></span> | [<span data-ttu-id="05a29-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="05a29-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="05a29-149">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="05a29-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="05a29-151">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-151">Compose</span></span><br><span data-ttu-id="05a29-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-152">Read</span></span> | [<span data-ttu-id="05a29-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="05a29-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="05a29-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="05a29-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="05a29-155">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="05a29-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="05a29-156">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-156">Compose</span></span><br><span data-ttu-id="05a29-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-157">Read</span></span> | [<span data-ttu-id="05a29-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="05a29-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="05a29-159">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-160">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="05a29-161">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-161">Compose</span></span><br><span data-ttu-id="05a29-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-162">Read</span></span> | [<span data-ttu-id="05a29-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="05a29-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="05a29-164">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="05a29-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="05a29-166">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-166">Compose</span></span><br><span data-ttu-id="05a29-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-167">Read</span></span> | [<span data-ttu-id="05a29-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="05a29-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="05a29-169">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="05a29-170">UI</span><span class="sxs-lookup"><span data-stu-id="05a29-170">ui</span></span>](#ui-ui) | <span data-ttu-id="05a29-171">作成</span><span class="sxs-lookup"><span data-stu-id="05a29-171">Compose</span></span><br><span data-ttu-id="05a29-172">読み取り</span><span class="sxs-lookup"><span data-stu-id="05a29-172">Read</span></span> | [<span data-ttu-id="05a29-173">UI</span><span class="sxs-lookup"><span data-stu-id="05a29-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="05a29-174">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="05a29-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="05a29-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="05a29-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="05a29-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="05a29-177">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office ホストがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="05a29-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="05a29-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="05a29-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-179">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-179">Type</span></span>

*   [<span data-ttu-id="05a29-180">Auth</span><span class="sxs-lookup"><span data-stu-id="05a29-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="05a29-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-181">Requirements</span></span>

|<span data-ttu-id="05a29-182">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-182">Requirement</span></span>| <span data-ttu-id="05a29-183">値</span><span class="sxs-lookup"><span data-stu-id="05a29-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="05a29-185">Preview</span></span>|
|[<span data-ttu-id="05a29-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-188">例</span><span class="sxs-lookup"><span data-stu-id="05a29-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="05a29-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="05a29-189">contentLanguage: String</span></span>

<span data-ttu-id="05a29-190">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="05a29-191">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="05a29-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-192">型</span><span class="sxs-lookup"><span data-stu-id="05a29-192">Type</span></span>

*   <span data-ttu-id="05a29-193">String</span><span class="sxs-lookup"><span data-stu-id="05a29-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="05a29-194">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-194">Requirements</span></span>

|<span data-ttu-id="05a29-195">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-195">Requirement</span></span>| <span data-ttu-id="05a29-196">値</span><span class="sxs-lookup"><span data-stu-id="05a29-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-198">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-198">1.1</span></span>|
|[<span data-ttu-id="05a29-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-201">例</span><span class="sxs-lookup"><span data-stu-id="05a29-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="05a29-202">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="05a29-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="05a29-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-204">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-204">Type</span></span>

*   [<span data-ttu-id="05a29-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="05a29-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="05a29-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-206">Requirements</span></span>

|<span data-ttu-id="05a29-207">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-207">Requirement</span></span>| <span data-ttu-id="05a29-208">値</span><span class="sxs-lookup"><span data-stu-id="05a29-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-210">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-210">1.1</span></span>|
|[<span data-ttu-id="05a29-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-213">例</span><span class="sxs-lookup"><span data-stu-id="05a29-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="05a29-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="05a29-214">displayLanguage: String</span></span>

<span data-ttu-id="05a29-215">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="05a29-216">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="05a29-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-217">型</span><span class="sxs-lookup"><span data-stu-id="05a29-217">Type</span></span>

*   <span data-ttu-id="05a29-218">文字列</span><span class="sxs-lookup"><span data-stu-id="05a29-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="05a29-219">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-219">Requirements</span></span>

|<span data-ttu-id="05a29-220">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-220">Requirement</span></span>| <span data-ttu-id="05a29-221">値</span><span class="sxs-lookup"><span data-stu-id="05a29-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-223">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-223">1.1</span></span>|
|[<span data-ttu-id="05a29-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-226">例</span><span class="sxs-lookup"><span data-stu-id="05a29-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="05a29-227">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="05a29-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="05a29-228">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-229">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-229">Type</span></span>

*   [<span data-ttu-id="05a29-230">HostType</span><span class="sxs-lookup"><span data-stu-id="05a29-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="05a29-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-231">Requirements</span></span>

|<span data-ttu-id="05a29-232">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-232">Requirement</span></span>| <span data-ttu-id="05a29-233">値</span><span class="sxs-lookup"><span data-stu-id="05a29-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-235">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-235">1.1</span></span>|
|[<span data-ttu-id="05a29-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-238">例</span><span class="sxs-lookup"><span data-stu-id="05a29-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="05a29-239">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="05a29-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="05a29-240">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="05a29-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="05a29-241">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="05a29-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="05a29-242">Office テーマの色を使用すると、アドインの配色を、[**ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office ホストアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="05a29-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="05a29-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="05a29-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-244">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-244">Type</span></span>

*   [<span data-ttu-id="05a29-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="05a29-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="05a29-246">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="05a29-246">Properties:</span></span>

|<span data-ttu-id="05a29-247">名前</span><span class="sxs-lookup"><span data-stu-id="05a29-247">Name</span></span>| <span data-ttu-id="05a29-248">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-248">Type</span></span>| <span data-ttu-id="05a29-249">説明</span><span class="sxs-lookup"><span data-stu-id="05a29-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="05a29-250">文字列</span><span class="sxs-lookup"><span data-stu-id="05a29-250">String</span></span>|<span data-ttu-id="05a29-251">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="05a29-252">文字列</span><span class="sxs-lookup"><span data-stu-id="05a29-252">String</span></span>|<span data-ttu-id="05a29-253">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="05a29-254">String</span><span class="sxs-lookup"><span data-stu-id="05a29-254">String</span></span>|<span data-ttu-id="05a29-255">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="05a29-256">String</span><span class="sxs-lookup"><span data-stu-id="05a29-256">String</span></span>|<span data-ttu-id="05a29-257">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05a29-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-258">Requirements</span></span>

|<span data-ttu-id="05a29-259">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-259">Requirement</span></span>| <span data-ttu-id="05a29-260">値</span><span class="sxs-lookup"><span data-stu-id="05a29-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-262">プレビュー</span><span class="sxs-lookup"><span data-stu-id="05a29-262">Preview</span></span>|
|[<span data-ttu-id="05a29-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-265">例</span><span class="sxs-lookup"><span data-stu-id="05a29-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="05a29-266">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="05a29-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="05a29-267">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="05a29-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-268">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-268">Type</span></span>

*   [<span data-ttu-id="05a29-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="05a29-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="05a29-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-270">Requirements</span></span>

|<span data-ttu-id="05a29-271">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-271">Requirement</span></span>| <span data-ttu-id="05a29-272">値</span><span class="sxs-lookup"><span data-stu-id="05a29-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-274">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-274">1.1</span></span>|
|[<span data-ttu-id="05a29-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-277">例</span><span class="sxs-lookup"><span data-stu-id="05a29-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="05a29-278">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="05a29-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="05a29-279">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="05a29-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-280">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-280">Type</span></span>

*   [<span data-ttu-id="05a29-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="05a29-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="05a29-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-282">Requirements</span></span>

|<span data-ttu-id="05a29-283">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-283">Requirement</span></span>| <span data-ttu-id="05a29-284">値</span><span class="sxs-lookup"><span data-stu-id="05a29-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-286">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-286">1.1</span></span>|
|[<span data-ttu-id="05a29-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-288">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="05a29-289">例</span><span class="sxs-lookup"><span data-stu-id="05a29-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="05a29-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="05a29-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="05a29-291">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="05a29-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="05a29-292">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="05a29-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-293">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-293">Type</span></span>

*   [<span data-ttu-id="05a29-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="05a29-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="05a29-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-295">Requirements</span></span>

|<span data-ttu-id="05a29-296">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-296">Requirement</span></span>| <span data-ttu-id="05a29-297">値</span><span class="sxs-lookup"><span data-stu-id="05a29-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-299">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-299">1.1</span></span>|
|[<span data-ttu-id="05a29-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="05a29-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="05a29-301">制限あり</span><span class="sxs-lookup"><span data-stu-id="05a29-301">Restricted</span></span>|
|[<span data-ttu-id="05a29-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-303">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="05a29-304">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="05a29-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="05a29-305">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="05a29-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="05a29-306">種類</span><span class="sxs-lookup"><span data-stu-id="05a29-306">Type</span></span>

*   [<span data-ttu-id="05a29-307">UI</span><span class="sxs-lookup"><span data-stu-id="05a29-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="05a29-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="05a29-308">Requirements</span></span>

|<span data-ttu-id="05a29-309">要件</span><span class="sxs-lookup"><span data-stu-id="05a29-309">Requirement</span></span>| <span data-ttu-id="05a29-310">値</span><span class="sxs-lookup"><span data-stu-id="05a29-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="05a29-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="05a29-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="05a29-312">1.1</span><span class="sxs-lookup"><span data-stu-id="05a29-312">1.1</span></span>|
|[<span data-ttu-id="05a29-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="05a29-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="05a29-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="05a29-314">Compose or Read</span></span>|
