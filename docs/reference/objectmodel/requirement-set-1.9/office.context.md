---
title: Office.context - 要件セット 1.9
description: Office。メールボックス API 要件セット 1.9 をOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590520"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="7e2ba-103">context (メールボックス要件セット 1.9)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="7e2ba-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="7e2ba-104">[Office](office.md).context</span></span>

<span data-ttu-id="7e2ba-105">Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="7e2ba-106">この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e2ba-107">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-107">Requirements</span></span>

|<span data-ttu-id="7e2ba-108">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-108">Requirement</span></span>| <span data-ttu-id="7e2ba-109">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-111">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-111">1.1</span></span>|
|[<span data-ttu-id="7e2ba-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="7e2ba-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7e2ba-114">Properties</span></span>

| <span data-ttu-id="7e2ba-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7e2ba-115">Property</span></span> | <span data-ttu-id="7e2ba-116">モード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-116">Modes</span></span> | <span data-ttu-id="7e2ba-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="7e2ba-117">Return type</span></span> | <span data-ttu-id="7e2ba-118">最小値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-118">Minimum</span></span><br><span data-ttu-id="7e2ba-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="7e2ba-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7e2ba-120">auth</span><span class="sxs-lookup"><span data-stu-id="7e2ba-120">auth</span></span>](#auth-auth) | <span data-ttu-id="7e2ba-121">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-121">Compose</span></span><br><span data-ttu-id="7e2ba-122">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-122">Read</span></span> | [<span data-ttu-id="7e2ba-123">Auth</span><span class="sxs-lookup"><span data-stu-id="7e2ba-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="7e2ba-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="7e2ba-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="7e2ba-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="7e2ba-126">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-126">Compose</span></span><br><span data-ttu-id="7e2ba-127">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-127">Read</span></span> | <span data-ttu-id="7e2ba-128">String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-128">String</span></span> | [<span data-ttu-id="7e2ba-129">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-130">診断</span><span class="sxs-lookup"><span data-stu-id="7e2ba-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="7e2ba-131">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-131">Compose</span></span><br><span data-ttu-id="7e2ba-132">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-132">Read</span></span> | [<span data-ttu-id="7e2ba-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7e2ba-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7e2ba-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7e2ba-136">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-136">Compose</span></span><br><span data-ttu-id="7e2ba-137">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-137">Read</span></span> | <span data-ttu-id="7e2ba-138">String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-138">String</span></span> | [<span data-ttu-id="7e2ba-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-140">host</span><span class="sxs-lookup"><span data-stu-id="7e2ba-140">host</span></span>](#host-hosttype) | <span data-ttu-id="7e2ba-141">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-141">Compose</span></span><br><span data-ttu-id="7e2ba-142">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-142">Read</span></span> | [<span data-ttu-id="7e2ba-143">HostType</span><span class="sxs-lookup"><span data-stu-id="7e2ba-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-144">1.5</span><span class="sxs-lookup"><span data-stu-id="7e2ba-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7e2ba-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="7e2ba-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="7e2ba-146">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-146">Compose</span></span><br><span data-ttu-id="7e2ba-147">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-147">Read</span></span> | [<span data-ttu-id="7e2ba-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="7e2ba-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-149">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-150">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7e2ba-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="7e2ba-151">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-151">Compose</span></span><br><span data-ttu-id="7e2ba-152">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-152">Read</span></span> | [<span data-ttu-id="7e2ba-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7e2ba-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-154">1.5</span><span class="sxs-lookup"><span data-stu-id="7e2ba-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7e2ba-155">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="7e2ba-156">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-156">Compose</span></span><br><span data-ttu-id="7e2ba-157">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-157">Read</span></span> | [<span data-ttu-id="7e2ba-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7e2ba-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-159">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e2ba-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7e2ba-161">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-161">Compose</span></span><br><span data-ttu-id="7e2ba-162">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-162">Read</span></span> | [<span data-ttu-id="7e2ba-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e2ba-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e2ba-165">UI</span><span class="sxs-lookup"><span data-stu-id="7e2ba-165">ui</span></span>](#ui-ui) | <span data-ttu-id="7e2ba-166">作成</span><span class="sxs-lookup"><span data-stu-id="7e2ba-166">Compose</span></span><br><span data-ttu-id="7e2ba-167">Read</span><span class="sxs-lookup"><span data-stu-id="7e2ba-167">Read</span></span> | [<span data-ttu-id="7e2ba-168">UI</span><span class="sxs-lookup"><span data-stu-id="7e2ba-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7e2ba-169">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="7e2ba-170">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="7e2ba-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="7e2ba-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="7e2ba-172">シングル[サインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするには、Office アプリケーションがアドインの Web アプリケーションへのアクセス トークンを取得できるメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="7e2ba-173">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="7e2ba-174">[「IdentityAPI 1.3 要件セット」を参照してください](../../requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-175">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-175">Type</span></span>

*   [<span data-ttu-id="7e2ba-176">Auth</span><span class="sxs-lookup"><span data-stu-id="7e2ba-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-177">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-177">Requirements</span></span>

|<span data-ttu-id="7e2ba-178">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-178">Requirement</span></span>| <span data-ttu-id="7e2ba-179">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-181">該当なし</span><span class="sxs-lookup"><span data-stu-id="7e2ba-181">N/A</span></span>|
|[<span data-ttu-id="7e2ba-182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-183">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-184">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="7e2ba-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-185">contentLanguage: String</span></span>

<span data-ttu-id="7e2ba-186">アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="7e2ba-187">この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 </span><span class="sxs-lookup"><span data-stu-id="7e2ba-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-188">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-188">Type</span></span>

*   <span data-ttu-id="7e2ba-189">String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e2ba-190">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-190">Requirements</span></span>

|<span data-ttu-id="7e2ba-191">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-191">Requirement</span></span>| <span data-ttu-id="7e2ba-192">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-194">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-194">1.1</span></span>|
|[<span data-ttu-id="7e2ba-195">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-196">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-197">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="7e2ba-198">診断: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="7e2ba-199">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-200">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-200">Type</span></span>

*   [<span data-ttu-id="7e2ba-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7e2ba-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-202">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-202">Requirements</span></span>

|<span data-ttu-id="7e2ba-203">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-203">Requirement</span></span>| <span data-ttu-id="7e2ba-204">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-206">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-206">1.1</span></span>|
|[<span data-ttu-id="7e2ba-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-209">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="7e2ba-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-210">displayLanguage: String</span></span>

<span data-ttu-id="7e2ba-211">ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="7e2ba-212">この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-213">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-213">Type</span></span>

*   <span data-ttu-id="7e2ba-214">String</span><span class="sxs-lookup"><span data-stu-id="7e2ba-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e2ba-215">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-215">Requirements</span></span>

|<span data-ttu-id="7e2ba-216">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-216">Requirement</span></span>| <span data-ttu-id="7e2ba-217">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-219">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-219">1.1</span></span>|
|[<span data-ttu-id="7e2ba-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-221">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-222">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="7e2ba-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="7e2ba-224">アドインをOfficeしているアプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7e2ba-225">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-226">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-226">Type</span></span>

*   [<span data-ttu-id="7e2ba-227">HostType</span><span class="sxs-lookup"><span data-stu-id="7e2ba-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-228">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-228">Requirements</span></span>

|<span data-ttu-id="7e2ba-229">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-229">Requirement</span></span>| <span data-ttu-id="7e2ba-230">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-231">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-232">1.5</span><span class="sxs-lookup"><span data-stu-id="7e2ba-232">1.5</span></span>|
|[<span data-ttu-id="7e2ba-233">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-234">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-235">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="7e2ba-236">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="7e2ba-237">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="7e2ba-238">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-239">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-239">Type</span></span>

*   [<span data-ttu-id="7e2ba-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7e2ba-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-241">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-241">Requirements</span></span>

|<span data-ttu-id="7e2ba-242">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-242">Requirement</span></span>| <span data-ttu-id="7e2ba-243">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-244">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-245">1.5</span><span class="sxs-lookup"><span data-stu-id="7e2ba-245">1.5</span></span>|
|[<span data-ttu-id="7e2ba-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-247">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-248">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="7e2ba-249">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="7e2ba-250">現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-251">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-251">Type</span></span>

*   [<span data-ttu-id="7e2ba-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7e2ba-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-253">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-253">Requirements</span></span>

|<span data-ttu-id="7e2ba-254">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-254">Requirement</span></span>| <span data-ttu-id="7e2ba-255">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-256">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-257">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-257">1.1</span></span>|
|[<span data-ttu-id="7e2ba-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-259">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e2ba-260">例</span><span class="sxs-lookup"><span data-stu-id="7e2ba-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="7e2ba-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="7e2ba-262">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7e2ba-263">このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`</span><span class="sxs-lookup"><span data-stu-id="7e2ba-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-264">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-264">Type</span></span>

*   [<span data-ttu-id="7e2ba-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e2ba-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-266">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-266">Requirements</span></span>

|<span data-ttu-id="7e2ba-267">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-267">Requirement</span></span>| <span data-ttu-id="7e2ba-268">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-269">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-270">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-270">1.1</span></span>|
|[<span data-ttu-id="7e2ba-271">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7e2ba-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="7e2ba-272">制限あり</span><span class="sxs-lookup"><span data-stu-id="7e2ba-272">Restricted</span></span>|
|[<span data-ttu-id="7e2ba-273">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-274">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="7e2ba-275">ui: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="7e2ba-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="7e2ba-276">ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。</span><span class="sxs-lookup"><span data-stu-id="7e2ba-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7e2ba-277">型</span><span class="sxs-lookup"><span data-stu-id="7e2ba-277">Type</span></span>

*   [<span data-ttu-id="7e2ba-278">UI</span><span class="sxs-lookup"><span data-stu-id="7e2ba-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="7e2ba-279">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-279">Requirements</span></span>

|<span data-ttu-id="7e2ba-280">要件</span><span class="sxs-lookup"><span data-stu-id="7e2ba-280">Requirement</span></span>| <span data-ttu-id="7e2ba-281">値</span><span class="sxs-lookup"><span data-stu-id="7e2ba-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e2ba-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7e2ba-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e2ba-283">1.1</span><span class="sxs-lookup"><span data-stu-id="7e2ba-283">1.1</span></span>|
|[<span data-ttu-id="7e2ba-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7e2ba-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e2ba-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7e2ba-285">Compose or Read</span></span>|
