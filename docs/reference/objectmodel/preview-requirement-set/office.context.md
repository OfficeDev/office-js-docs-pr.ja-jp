---
title: Office.context - プレビュー要件セット
description: Office。メールボックス API プレビュー要件セットをOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 59b1cce579afe69384e41a6f31cc70c8cec25bea
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591073"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="e2d83-103">context (メールボックス プレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="e2d83-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="e2d83-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e2d83-104">[Office](office.md).context</span></span>

<span data-ttu-id="e2d83-105">Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="e2d83-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e2d83-106">この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="e2d83-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d83-107">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-107">Requirements</span></span>

|<span data-ttu-id="e2d83-108">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-108">Requirement</span></span>| <span data-ttu-id="e2d83-109">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-111">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-111">1.1</span></span>|
|[<span data-ttu-id="e2d83-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="e2d83-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e2d83-114">Properties</span></span>

| <span data-ttu-id="e2d83-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e2d83-115">Property</span></span> | <span data-ttu-id="e2d83-116">モード</span><span class="sxs-lookup"><span data-stu-id="e2d83-116">Modes</span></span> | <span data-ttu-id="e2d83-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="e2d83-117">Return type</span></span> | <span data-ttu-id="e2d83-118">最小値</span><span class="sxs-lookup"><span data-stu-id="e2d83-118">Minimum</span></span><br><span data-ttu-id="e2d83-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="e2d83-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e2d83-120">auth</span><span class="sxs-lookup"><span data-stu-id="e2d83-120">auth</span></span>](#auth-auth) | <span data-ttu-id="e2d83-121">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-121">Compose</span></span><br><span data-ttu-id="e2d83-122">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-122">Read</span></span> | [<span data-ttu-id="e2d83-123">Auth</span><span class="sxs-lookup"><span data-stu-id="e2d83-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="e2d83-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="e2d83-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e2d83-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e2d83-126">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-126">Compose</span></span><br><span data-ttu-id="e2d83-127">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-127">Read</span></span> | <span data-ttu-id="e2d83-128">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-128">String</span></span> | [<span data-ttu-id="e2d83-129">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-130">診断</span><span class="sxs-lookup"><span data-stu-id="e2d83-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e2d83-131">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-131">Compose</span></span><br><span data-ttu-id="e2d83-132">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-132">Read</span></span> | [<span data-ttu-id="e2d83-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e2d83-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e2d83-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e2d83-136">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-136">Compose</span></span><br><span data-ttu-id="e2d83-137">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-137">Read</span></span> | <span data-ttu-id="e2d83-138">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-138">String</span></span> | [<span data-ttu-id="e2d83-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-140">host</span><span class="sxs-lookup"><span data-stu-id="e2d83-140">host</span></span>](#host-hosttype) | <span data-ttu-id="e2d83-141">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-141">Compose</span></span><br><span data-ttu-id="e2d83-142">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-142">Read</span></span> | [<span data-ttu-id="e2d83-143">HostType</span><span class="sxs-lookup"><span data-stu-id="e2d83-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-144">1.5</span><span class="sxs-lookup"><span data-stu-id="e2d83-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e2d83-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="e2d83-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e2d83-146">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-146">Compose</span></span><br><span data-ttu-id="e2d83-147">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-147">Read</span></span> | [<span data-ttu-id="e2d83-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="e2d83-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="e2d83-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="e2d83-151">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-151">Compose</span></span><br><span data-ttu-id="e2d83-152">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-152">Read</span></span> | [<span data-ttu-id="e2d83-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="e2d83-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e2d83-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="e2d83-155">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e2d83-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="e2d83-156">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-156">Compose</span></span><br><span data-ttu-id="e2d83-157">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-157">Read</span></span> | [<span data-ttu-id="e2d83-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e2d83-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-159">1.5</span><span class="sxs-lookup"><span data-stu-id="e2d83-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e2d83-160">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e2d83-161">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-161">Compose</span></span><br><span data-ttu-id="e2d83-162">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-162">Read</span></span> | [<span data-ttu-id="e2d83-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e2d83-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-164">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e2d83-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e2d83-166">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-166">Compose</span></span><br><span data-ttu-id="e2d83-167">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-167">Read</span></span> | [<span data-ttu-id="e2d83-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e2d83-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e2d83-170">UI</span><span class="sxs-lookup"><span data-stu-id="e2d83-170">ui</span></span>](#ui-ui) | <span data-ttu-id="e2d83-171">作成</span><span class="sxs-lookup"><span data-stu-id="e2d83-171">Compose</span></span><br><span data-ttu-id="e2d83-172">Read</span><span class="sxs-lookup"><span data-stu-id="e2d83-172">Read</span></span> | [<span data-ttu-id="e2d83-173">UI</span><span class="sxs-lookup"><span data-stu-id="e2d83-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e2d83-174">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e2d83-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="e2d83-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="e2d83-176">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="e2d83-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="e2d83-177">シングル[サインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするには、Office アプリケーションがアドインの Web アプリケーションへのアクセス トークンを取得できるメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="e2d83-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="e2d83-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-179">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-179">Type</span></span>

*   [<span data-ttu-id="e2d83-180">Auth</span><span class="sxs-lookup"><span data-stu-id="e2d83-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="e2d83-181">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-181">Requirements</span></span>

|<span data-ttu-id="e2d83-182">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-182">Requirement</span></span>| <span data-ttu-id="e2d83-183">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e2d83-185">Preview</span></span>|
|[<span data-ttu-id="e2d83-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-188">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="e2d83-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="e2d83-189">contentLanguage: String</span></span>

<span data-ttu-id="e2d83-190">アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e2d83-191">この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 </span><span class="sxs-lookup"><span data-stu-id="e2d83-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-192">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-192">Type</span></span>

*   <span data-ttu-id="e2d83-193">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d83-194">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-194">Requirements</span></span>

|<span data-ttu-id="e2d83-195">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-195">Requirement</span></span>| <span data-ttu-id="e2d83-196">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-198">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-198">1.1</span></span>|
|[<span data-ttu-id="e2d83-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-201">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="e2d83-202">診断: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e2d83-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e2d83-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-204">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-204">Type</span></span>

*   [<span data-ttu-id="e2d83-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e2d83-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e2d83-206">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-206">Requirements</span></span>

|<span data-ttu-id="e2d83-207">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-207">Requirement</span></span>| <span data-ttu-id="e2d83-208">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-210">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-210">1.1</span></span>|
|[<span data-ttu-id="e2d83-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-213">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e2d83-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="e2d83-214">displayLanguage: String</span></span>

<span data-ttu-id="e2d83-215">ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="e2d83-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="e2d83-216">この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。</span><span class="sxs-lookup"><span data-stu-id="e2d83-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-217">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-217">Type</span></span>

*   <span data-ttu-id="e2d83-218">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2d83-219">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-219">Requirements</span></span>

|<span data-ttu-id="e2d83-220">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-220">Requirement</span></span>| <span data-ttu-id="e2d83-221">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-223">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-223">1.1</span></span>|
|[<span data-ttu-id="e2d83-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-226">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="e2d83-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="e2d83-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="e2d83-228">アドインをOfficeしているアプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d83-229">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してホストを取得できます。</span><span class="sxs-lookup"><span data-stu-id="e2d83-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-230">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-230">Type</span></span>

*   [<span data-ttu-id="e2d83-231">HostType</span><span class="sxs-lookup"><span data-stu-id="e2d83-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="e2d83-232">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-232">Requirements</span></span>

|<span data-ttu-id="e2d83-233">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-233">Requirement</span></span>| <span data-ttu-id="e2d83-234">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-236">1.5</span><span class="sxs-lookup"><span data-stu-id="e2d83-236">1.5</span></span>|
|[<span data-ttu-id="e2d83-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-239">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="e2d83-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="e2d83-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="e2d83-241">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="e2d83-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d83-242">このメンバーは、このメンバーのOutlookでのみWindows。</span><span class="sxs-lookup"><span data-stu-id="e2d83-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="e2d83-243">Office テーマの色を使用すると、すべての Office クライアント アプリケーションに適用される File **> Office Account > Office** テーマ UI を使用して、ユーザーが選択した現在の Office テーマとアドインの配色を調整できます。</span><span class="sxs-lookup"><span data-stu-id="e2d83-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="e2d83-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="e2d83-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-245">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-245">Type</span></span>

*   [<span data-ttu-id="e2d83-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="e2d83-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="e2d83-247">プロパティ</span><span class="sxs-lookup"><span data-stu-id="e2d83-247">Properties</span></span>

|<span data-ttu-id="e2d83-248">名前</span><span class="sxs-lookup"><span data-stu-id="e2d83-248">Name</span></span>| <span data-ttu-id="e2d83-249">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-249">Type</span></span>| <span data-ttu-id="e2d83-250">説明</span><span class="sxs-lookup"><span data-stu-id="e2d83-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="e2d83-251">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-251">String</span></span>|<span data-ttu-id="e2d83-252">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="e2d83-253">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-253">String</span></span>|<span data-ttu-id="e2d83-254">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="e2d83-255">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-255">String</span></span>|<span data-ttu-id="e2d83-256">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="e2d83-257">String</span><span class="sxs-lookup"><span data-stu-id="e2d83-257">String</span></span>|<span data-ttu-id="e2d83-258">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e2d83-259">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-259">Requirements</span></span>

|<span data-ttu-id="e2d83-260">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-260">Requirement</span></span>| <span data-ttu-id="e2d83-261">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-262">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-263">プレビュー</span><span class="sxs-lookup"><span data-stu-id="e2d83-263">Preview</span></span>|
|[<span data-ttu-id="e2d83-264">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-265">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-266">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="e2d83-267">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="e2d83-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="e2d83-268">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d83-269">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。</span><span class="sxs-lookup"><span data-stu-id="e2d83-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-270">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-270">Type</span></span>

*   [<span data-ttu-id="e2d83-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e2d83-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="e2d83-272">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-272">Requirements</span></span>

|<span data-ttu-id="e2d83-273">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-273">Requirement</span></span>| <span data-ttu-id="e2d83-274">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-275">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-276">1.5</span><span class="sxs-lookup"><span data-stu-id="e2d83-276">1.5</span></span>|
|[<span data-ttu-id="e2d83-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-278">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-279">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="e2d83-280">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e2d83-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e2d83-281">現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-282">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-282">Type</span></span>

*   [<span data-ttu-id="e2d83-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e2d83-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e2d83-284">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-284">Requirements</span></span>

|<span data-ttu-id="e2d83-285">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-285">Requirement</span></span>| <span data-ttu-id="e2d83-286">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-288">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-288">1.1</span></span>|
|[<span data-ttu-id="e2d83-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-290">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e2d83-291">例</span><span class="sxs-lookup"><span data-stu-id="e2d83-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="e2d83-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e2d83-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e2d83-293">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e2d83-294">このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`</span><span class="sxs-lookup"><span data-stu-id="e2d83-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-295">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-295">Type</span></span>

*   [<span data-ttu-id="e2d83-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e2d83-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e2d83-297">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-297">Requirements</span></span>

|<span data-ttu-id="e2d83-298">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-298">Requirement</span></span>| <span data-ttu-id="e2d83-299">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-300">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-301">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-301">1.1</span></span>|
|[<span data-ttu-id="e2d83-302">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="e2d83-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="e2d83-303">制限あり</span><span class="sxs-lookup"><span data-stu-id="e2d83-303">Restricted</span></span>|
|[<span data-ttu-id="e2d83-304">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-305">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="e2d83-306">ui: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e2d83-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e2d83-307">ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。</span><span class="sxs-lookup"><span data-stu-id="e2d83-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e2d83-308">型</span><span class="sxs-lookup"><span data-stu-id="e2d83-308">Type</span></span>

*   [<span data-ttu-id="e2d83-309">UI</span><span class="sxs-lookup"><span data-stu-id="e2d83-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e2d83-310">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-310">Requirements</span></span>

|<span data-ttu-id="e2d83-311">要件</span><span class="sxs-lookup"><span data-stu-id="e2d83-311">Requirement</span></span>| <span data-ttu-id="e2d83-312">値</span><span class="sxs-lookup"><span data-stu-id="e2d83-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2d83-313">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="e2d83-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e2d83-314">1.1</span><span class="sxs-lookup"><span data-stu-id="e2d83-314">1.1</span></span>|
|[<span data-ttu-id="e2d83-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="e2d83-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e2d83-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="e2d83-316">Compose or Read</span></span>|
