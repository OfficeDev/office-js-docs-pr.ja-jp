---
title: Office コンテキスト-プレビュー要件セット
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165518"
---
# <a name="context"></a><span data-ttu-id="c5dc2-102">context</span><span class="sxs-lookup"><span data-stu-id="c5dc2-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="c5dc2-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c5dc2-103">[Office](office.md).context</span></span>

<span data-ttu-id="c5dc2-104">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c5dc2-105">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5dc2-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-106">Requirements</span></span>

|<span data-ttu-id="c5dc2-107">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-107">Requirement</span></span>| <span data-ttu-id="c5dc2-108">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-110">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-110">1.1</span></span>|
|[<span data-ttu-id="c5dc2-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c5dc2-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="c5dc2-113">Properties</span></span>

| <span data-ttu-id="c5dc2-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="c5dc2-114">Property</span></span> | <span data-ttu-id="c5dc2-115">モード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-115">Modes</span></span> | <span data-ttu-id="c5dc2-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="c5dc2-116">Return type</span></span> | <span data-ttu-id="c5dc2-117">最小値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-117">Minimum</span></span><br><span data-ttu-id="c5dc2-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="c5dc2-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c5dc2-119">authoritative</span><span class="sxs-lookup"><span data-stu-id="c5dc2-119">auth</span></span>](#auth-auth) | <span data-ttu-id="c5dc2-120">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-120">Compose</span></span><br><span data-ttu-id="c5dc2-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-121">Read</span></span> | [<span data-ttu-id="c5dc2-122">Auth</span><span class="sxs-lookup"><span data-stu-id="c5dc2-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-123">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c5dc2-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="c5dc2-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c5dc2-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c5dc2-125">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-125">Compose</span></span><br><span data-ttu-id="c5dc2-126">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-126">Read</span></span> | <span data-ttu-id="c5dc2-127">文字列</span><span class="sxs-lookup"><span data-stu-id="c5dc2-127">String</span></span> | [<span data-ttu-id="c5dc2-128">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-129">ダン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c5dc2-130">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-130">Compose</span></span><br><span data-ttu-id="c5dc2-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-131">Read</span></span> | [<span data-ttu-id="c5dc2-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5dc2-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-133">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c5dc2-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c5dc2-135">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-135">Compose</span></span><br><span data-ttu-id="c5dc2-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-136">Read</span></span> | <span data-ttu-id="c5dc2-137">文字列</span><span class="sxs-lookup"><span data-stu-id="c5dc2-137">String</span></span> | [<span data-ttu-id="c5dc2-138">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-139">主催</span><span class="sxs-lookup"><span data-stu-id="c5dc2-139">host</span></span>](#host-hosttype) | <span data-ttu-id="c5dc2-140">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-140">Compose</span></span><br><span data-ttu-id="c5dc2-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-141">Read</span></span> | [<span data-ttu-id="c5dc2-142">HostType</span><span class="sxs-lookup"><span data-stu-id="c5dc2-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-143">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="c5dc2-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c5dc2-145">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-145">Compose</span></span><br><span data-ttu-id="c5dc2-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-146">Read</span></span> | [<span data-ttu-id="c5dc2-147">メールボックス</span><span class="sxs-lookup"><span data-stu-id="c5dc2-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-148">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="c5dc2-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="c5dc2-150">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-150">Compose</span></span><br><span data-ttu-id="c5dc2-151">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-151">Read</span></span> | [<span data-ttu-id="c5dc2-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="c5dc2-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-153">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c5dc2-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="c5dc2-154">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c5dc2-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c5dc2-155">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-155">Compose</span></span><br><span data-ttu-id="c5dc2-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-156">Read</span></span> | [<span data-ttu-id="c5dc2-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5dc2-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-158">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-159">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c5dc2-160">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-160">Compose</span></span><br><span data-ttu-id="c5dc2-161">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-161">Read</span></span> | [<span data-ttu-id="c5dc2-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5dc2-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-163">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5dc2-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c5dc2-165">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-165">Compose</span></span><br><span data-ttu-id="c5dc2-166">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-166">Read</span></span> | [<span data-ttu-id="c5dc2-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5dc2-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-168">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c5dc2-169">UI</span><span class="sxs-lookup"><span data-stu-id="c5dc2-169">ui</span></span>](#ui-ui) | <span data-ttu-id="c5dc2-170">作成</span><span class="sxs-lookup"><span data-stu-id="c5dc2-170">Compose</span></span><br><span data-ttu-id="c5dc2-171">読み取り</span><span class="sxs-lookup"><span data-stu-id="c5dc2-171">Read</span></span> | [<span data-ttu-id="c5dc2-172">UI</span><span class="sxs-lookup"><span data-stu-id="c5dc2-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="c5dc2-173">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c5dc2-174">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="c5dc2-174">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="c5dc2-175">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="c5dc2-176">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office ホストがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-176">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="c5dc2-177">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-178">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-178">Type</span></span>

*   [<span data-ttu-id="c5dc2-179">Auth</span><span class="sxs-lookup"><span data-stu-id="c5dc2-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-180">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-180">Requirements</span></span>

|<span data-ttu-id="c5dc2-181">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-181">Requirement</span></span>| <span data-ttu-id="c5dc2-182">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-183">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-184">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c5dc2-184">Preview</span></span>|
|[<span data-ttu-id="c5dc2-185">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-185">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-186">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-187">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="c5dc2-188">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="c5dc2-188">contentLanguage: String</span></span>

<span data-ttu-id="c5dc2-189">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c5dc2-190">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-191">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-191">Type</span></span>

*   <span data-ttu-id="c5dc2-192">String</span><span class="sxs-lookup"><span data-stu-id="c5dc2-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5dc2-193">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-193">Requirements</span></span>

|<span data-ttu-id="c5dc2-194">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-194">Requirement</span></span>| <span data-ttu-id="c5dc2-195">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-196">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-197">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-197">1.1</span></span>|
|[<span data-ttu-id="c5dc2-198">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-198">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-199">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-200">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-200">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c5dc2-201">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c5dc2-202">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-203">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-203">Type</span></span>

*   [<span data-ttu-id="c5dc2-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="c5dc2-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-205">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-205">Requirements</span></span>

|<span data-ttu-id="c5dc2-206">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-206">Requirement</span></span>| <span data-ttu-id="c5dc2-207">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-208">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-209">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-209">1.1</span></span>|
|[<span data-ttu-id="c5dc2-210">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-210">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-211">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-212">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c5dc2-213">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="c5dc2-213">displayLanguage: String</span></span>

<span data-ttu-id="c5dc2-214">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c5dc2-215">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-216">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-216">Type</span></span>

*   <span data-ttu-id="c5dc2-217">文字列</span><span class="sxs-lookup"><span data-stu-id="c5dc2-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5dc2-218">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-218">Requirements</span></span>

|<span data-ttu-id="c5dc2-219">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-219">Requirement</span></span>| <span data-ttu-id="c5dc2-220">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-221">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-222">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-222">1.1</span></span>|
|[<span data-ttu-id="c5dc2-223">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-224">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-225">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-225">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c5dc2-226">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c5dc2-227">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-228">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-228">Type</span></span>

*   [<span data-ttu-id="c5dc2-229">HostType</span><span class="sxs-lookup"><span data-stu-id="c5dc2-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-230">Requirements</span></span>

|<span data-ttu-id="c5dc2-231">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-231">Requirement</span></span>| <span data-ttu-id="c5dc2-232">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-233">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-234">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-234">1.1</span></span>|
|[<span data-ttu-id="c5dc2-235">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-236">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-237">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="c5dc2-238">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="c5dc2-239">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c5dc2-240">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="c5dc2-241">Office テーマの色を使用すると、アドインの配色を、[**ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office ホストアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="c5dc2-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="c5dc2-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-243">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-243">Type</span></span>

*   [<span data-ttu-id="c5dc2-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="c5dc2-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="c5dc2-245">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="c5dc2-245">Properties:</span></span>

|<span data-ttu-id="c5dc2-246">名前</span><span class="sxs-lookup"><span data-stu-id="c5dc2-246">Name</span></span>| <span data-ttu-id="c5dc2-247">種類</span><span class="sxs-lookup"><span data-stu-id="c5dc2-247">Type</span></span>| <span data-ttu-id="c5dc2-248">説明</span><span class="sxs-lookup"><span data-stu-id="c5dc2-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c5dc2-249">文字列</span><span class="sxs-lookup"><span data-stu-id="c5dc2-249">String</span></span>|<span data-ttu-id="c5dc2-250">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c5dc2-251">文字列</span><span class="sxs-lookup"><span data-stu-id="c5dc2-251">String</span></span>|<span data-ttu-id="c5dc2-252">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c5dc2-253">String</span><span class="sxs-lookup"><span data-stu-id="c5dc2-253">String</span></span>|<span data-ttu-id="c5dc2-254">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c5dc2-255">String</span><span class="sxs-lookup"><span data-stu-id="c5dc2-255">String</span></span>|<span data-ttu-id="c5dc2-256">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c5dc2-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-257">Requirements</span></span>

|<span data-ttu-id="c5dc2-258">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-258">Requirement</span></span>| <span data-ttu-id="c5dc2-259">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-261">プレビュー</span><span class="sxs-lookup"><span data-stu-id="c5dc2-261">Preview</span></span>|
|[<span data-ttu-id="c5dc2-262">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-263">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-264">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-264">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="c5dc2-265">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c5dc2-266">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-267">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-267">Type</span></span>

*   [<span data-ttu-id="c5dc2-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c5dc2-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-269">Requirements</span></span>

|<span data-ttu-id="c5dc2-270">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-270">Requirement</span></span>| <span data-ttu-id="c5dc2-271">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-272">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-273">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-273">1.1</span></span>|
|[<span data-ttu-id="c5dc2-274">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-274">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-275">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-276">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c5dc2-277">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c5dc2-278">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-279">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-279">Type</span></span>

*   [<span data-ttu-id="c5dc2-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="c5dc2-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-281">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-281">Requirements</span></span>

|<span data-ttu-id="c5dc2-282">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-282">Requirement</span></span>| <span data-ttu-id="c5dc2-283">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-284">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-285">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-285">1.1</span></span>|
|[<span data-ttu-id="c5dc2-286">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-286">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-287">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c5dc2-288">例</span><span class="sxs-lookup"><span data-stu-id="c5dc2-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c5dc2-289">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c5dc2-290">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c5dc2-291">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-292">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-292">Type</span></span>

*   [<span data-ttu-id="c5dc2-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c5dc2-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-294">Requirements</span></span>

|<span data-ttu-id="c5dc2-295">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-295">Requirement</span></span>| <span data-ttu-id="c5dc2-296">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-297">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-298">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-298">1.1</span></span>|
|[<span data-ttu-id="c5dc2-299">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="c5dc2-299">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c5dc2-300">制限あり</span><span class="sxs-lookup"><span data-stu-id="c5dc2-300">Restricted</span></span>|
|[<span data-ttu-id="c5dc2-301">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-301">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-302">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c5dc2-303">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c5dc2-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c5dc2-304">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="c5dc2-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c5dc2-305">型</span><span class="sxs-lookup"><span data-stu-id="c5dc2-305">Type</span></span>

*   [<span data-ttu-id="c5dc2-306">UI</span><span class="sxs-lookup"><span data-stu-id="c5dc2-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c5dc2-307">Requirements</span><span class="sxs-lookup"><span data-stu-id="c5dc2-307">Requirements</span></span>

|<span data-ttu-id="c5dc2-308">要件</span><span class="sxs-lookup"><span data-stu-id="c5dc2-308">Requirement</span></span>| <span data-ttu-id="c5dc2-309">値</span><span class="sxs-lookup"><span data-stu-id="c5dc2-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5dc2-310">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="c5dc2-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c5dc2-311">1.1</span><span class="sxs-lookup"><span data-stu-id="c5dc2-311">1.1</span></span>|
|[<span data-ttu-id="c5dc2-312">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="c5dc2-312">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c5dc2-313">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="c5dc2-313">Compose or Read</span></span>|
