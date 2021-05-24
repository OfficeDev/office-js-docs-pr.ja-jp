---
title: Office.context - 要件セット 1.10
description: Office。メールボックス API 要件セット 1.10 をOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592063"
---
# <a name="context-mailbox-requirement-set-110"></a><span data-ttu-id="eb503-103">context (メールボックス要件セット 1.10)</span><span class="sxs-lookup"><span data-stu-id="eb503-103">context (Mailbox requirement set 1.10)</span></span>

### <a name="officecontext"></a><span data-ttu-id="eb503-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="eb503-104">[Office](office.md).context</span></span>

<span data-ttu-id="eb503-105">Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="eb503-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="eb503-106">この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="eb503-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb503-107">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-107">Requirements</span></span>

|<span data-ttu-id="eb503-108">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-108">Requirement</span></span>| <span data-ttu-id="eb503-109">値</span><span class="sxs-lookup"><span data-stu-id="eb503-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-111">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-111">1.1</span></span>|
|[<span data-ttu-id="eb503-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="eb503-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="eb503-114">Properties</span></span>

| <span data-ttu-id="eb503-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="eb503-115">Property</span></span> | <span data-ttu-id="eb503-116">モード</span><span class="sxs-lookup"><span data-stu-id="eb503-116">Modes</span></span> | <span data-ttu-id="eb503-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="eb503-117">Return type</span></span> | <span data-ttu-id="eb503-118">最小値</span><span class="sxs-lookup"><span data-stu-id="eb503-118">Minimum</span></span><br><span data-ttu-id="eb503-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="eb503-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="eb503-120">auth</span><span class="sxs-lookup"><span data-stu-id="eb503-120">auth</span></span>](#auth-auth) | <span data-ttu-id="eb503-121">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-121">Compose</span></span><br><span data-ttu-id="eb503-122">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-122">Read</span></span> | [<span data-ttu-id="eb503-123">Auth</span><span class="sxs-lookup"><span data-stu-id="eb503-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="eb503-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="eb503-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="eb503-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="eb503-126">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-126">Compose</span></span><br><span data-ttu-id="eb503-127">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-127">Read</span></span> | <span data-ttu-id="eb503-128">String</span><span class="sxs-lookup"><span data-stu-id="eb503-128">String</span></span> | [<span data-ttu-id="eb503-129">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-130">診断</span><span class="sxs-lookup"><span data-stu-id="eb503-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="eb503-131">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-131">Compose</span></span><br><span data-ttu-id="eb503-132">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-132">Read</span></span> | [<span data-ttu-id="eb503-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="eb503-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-134">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="eb503-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="eb503-136">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-136">Compose</span></span><br><span data-ttu-id="eb503-137">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-137">Read</span></span> | <span data-ttu-id="eb503-138">String</span><span class="sxs-lookup"><span data-stu-id="eb503-138">String</span></span> | [<span data-ttu-id="eb503-139">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-140">host</span><span class="sxs-lookup"><span data-stu-id="eb503-140">host</span></span>](#host-hosttype) | <span data-ttu-id="eb503-141">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-141">Compose</span></span><br><span data-ttu-id="eb503-142">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-142">Read</span></span> | [<span data-ttu-id="eb503-143">HostType</span><span class="sxs-lookup"><span data-stu-id="eb503-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-144">1.5</span><span class="sxs-lookup"><span data-stu-id="eb503-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="eb503-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="eb503-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="eb503-146">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-146">Compose</span></span><br><span data-ttu-id="eb503-147">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-147">Read</span></span> | [<span data-ttu-id="eb503-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="eb503-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-149">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-150">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="eb503-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="eb503-151">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-151">Compose</span></span><br><span data-ttu-id="eb503-152">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-152">Read</span></span> | [<span data-ttu-id="eb503-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="eb503-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-154">1.5</span><span class="sxs-lookup"><span data-stu-id="eb503-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="eb503-155">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="eb503-156">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-156">Compose</span></span><br><span data-ttu-id="eb503-157">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-157">Read</span></span> | [<span data-ttu-id="eb503-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="eb503-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-159">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb503-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="eb503-161">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-161">Compose</span></span><br><span data-ttu-id="eb503-162">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-162">Read</span></span> | [<span data-ttu-id="eb503-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb503-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-164">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb503-165">UI</span><span class="sxs-lookup"><span data-stu-id="eb503-165">ui</span></span>](#ui-ui) | <span data-ttu-id="eb503-166">作成</span><span class="sxs-lookup"><span data-stu-id="eb503-166">Compose</span></span><br><span data-ttu-id="eb503-167">Read</span><span class="sxs-lookup"><span data-stu-id="eb503-167">Read</span></span> | [<span data-ttu-id="eb503-168">UI</span><span class="sxs-lookup"><span data-stu-id="eb503-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="eb503-169">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="eb503-170">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="eb503-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="eb503-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="eb503-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="eb503-172">シングル[サインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするには、Office アプリケーションがアドインの Web アプリケーションへのアクセス トークンを取得できるメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="eb503-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="eb503-173">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="eb503-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-174">型</span><span class="sxs-lookup"><span data-stu-id="eb503-174">Type</span></span>

*   [<span data-ttu-id="eb503-175">Auth</span><span class="sxs-lookup"><span data-stu-id="eb503-175">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="eb503-176">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-176">Requirements</span></span>

|<span data-ttu-id="eb503-177">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-177">Requirement</span></span>| <span data-ttu-id="eb503-178">値</span><span class="sxs-lookup"><span data-stu-id="eb503-178">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-179">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-179">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-180">1.10</span><span class="sxs-lookup"><span data-stu-id="eb503-180">1.10</span></span>|
|[<span data-ttu-id="eb503-181">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-181">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-182">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-182">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-183">例</span><span class="sxs-lookup"><span data-stu-id="eb503-183">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="eb503-184">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="eb503-184">contentLanguage: String</span></span>

<span data-ttu-id="eb503-185">アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="eb503-185">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="eb503-186">この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 </span><span class="sxs-lookup"><span data-stu-id="eb503-186">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-187">型</span><span class="sxs-lookup"><span data-stu-id="eb503-187">Type</span></span>

*   <span data-ttu-id="eb503-188">String</span><span class="sxs-lookup"><span data-stu-id="eb503-188">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb503-189">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-189">Requirements</span></span>

|<span data-ttu-id="eb503-190">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-190">Requirement</span></span>| <span data-ttu-id="eb503-191">値</span><span class="sxs-lookup"><span data-stu-id="eb503-191">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-192">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-192">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-193">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-193">1.1</span></span>|
|[<span data-ttu-id="eb503-194">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-194">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-195">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-195">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-196">例</span><span class="sxs-lookup"><span data-stu-id="eb503-196">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="eb503-197">診断: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="eb503-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="eb503-198">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="eb503-198">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-199">型</span><span class="sxs-lookup"><span data-stu-id="eb503-199">Type</span></span>

*   [<span data-ttu-id="eb503-200">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="eb503-200">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="eb503-201">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-201">Requirements</span></span>

|<span data-ttu-id="eb503-202">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-202">Requirement</span></span>| <span data-ttu-id="eb503-203">値</span><span class="sxs-lookup"><span data-stu-id="eb503-203">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-204">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-204">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-205">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-205">1.1</span></span>|
|[<span data-ttu-id="eb503-206">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-206">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-207">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-207">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-208">例</span><span class="sxs-lookup"><span data-stu-id="eb503-208">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="eb503-209">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="eb503-209">displayLanguage: String</span></span>

<span data-ttu-id="eb503-210">ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="eb503-210">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="eb503-211">この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。</span><span class="sxs-lookup"><span data-stu-id="eb503-211">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-212">型</span><span class="sxs-lookup"><span data-stu-id="eb503-212">Type</span></span>

*   <span data-ttu-id="eb503-213">String</span><span class="sxs-lookup"><span data-stu-id="eb503-213">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb503-214">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-214">Requirements</span></span>

|<span data-ttu-id="eb503-215">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-215">Requirement</span></span>| <span data-ttu-id="eb503-216">値</span><span class="sxs-lookup"><span data-stu-id="eb503-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-217">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-217">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-218">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-218">1.1</span></span>|
|[<span data-ttu-id="eb503-219">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-219">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-220">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-221">例</span><span class="sxs-lookup"><span data-stu-id="eb503-221">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="eb503-222">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="eb503-222">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="eb503-223">アドインをOfficeしているアプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="eb503-223">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="eb503-224">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してホストを取得できます。</span><span class="sxs-lookup"><span data-stu-id="eb503-224">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-225">型</span><span class="sxs-lookup"><span data-stu-id="eb503-225">Type</span></span>

*   [<span data-ttu-id="eb503-226">HostType</span><span class="sxs-lookup"><span data-stu-id="eb503-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="eb503-227">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-227">Requirements</span></span>

|<span data-ttu-id="eb503-228">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-228">Requirement</span></span>| <span data-ttu-id="eb503-229">値</span><span class="sxs-lookup"><span data-stu-id="eb503-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-230">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-231">1.5</span><span class="sxs-lookup"><span data-stu-id="eb503-231">1.5</span></span>|
|[<span data-ttu-id="eb503-232">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-233">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-234">例</span><span class="sxs-lookup"><span data-stu-id="eb503-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="eb503-235">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="eb503-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="eb503-236">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="eb503-236">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="eb503-237">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。</span><span class="sxs-lookup"><span data-stu-id="eb503-237">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-238">型</span><span class="sxs-lookup"><span data-stu-id="eb503-238">Type</span></span>

*   [<span data-ttu-id="eb503-239">PlatformType</span><span class="sxs-lookup"><span data-stu-id="eb503-239">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="eb503-240">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-240">Requirements</span></span>

|<span data-ttu-id="eb503-241">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-241">Requirement</span></span>| <span data-ttu-id="eb503-242">値</span><span class="sxs-lookup"><span data-stu-id="eb503-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-243">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-244">1.5</span><span class="sxs-lookup"><span data-stu-id="eb503-244">1.5</span></span>|
|[<span data-ttu-id="eb503-245">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-246">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-247">例</span><span class="sxs-lookup"><span data-stu-id="eb503-247">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="eb503-248">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="eb503-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="eb503-249">現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="eb503-249">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-250">型</span><span class="sxs-lookup"><span data-stu-id="eb503-250">Type</span></span>

*   [<span data-ttu-id="eb503-251">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="eb503-251">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="eb503-252">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-252">Requirements</span></span>

|<span data-ttu-id="eb503-253">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-253">Requirement</span></span>| <span data-ttu-id="eb503-254">値</span><span class="sxs-lookup"><span data-stu-id="eb503-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-255">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-255">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-256">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-256">1.1</span></span>|
|[<span data-ttu-id="eb503-257">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-257">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-258">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb503-259">例</span><span class="sxs-lookup"><span data-stu-id="eb503-259">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="eb503-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="eb503-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="eb503-261">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="eb503-261">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="eb503-262">このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`</span><span class="sxs-lookup"><span data-stu-id="eb503-262">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-263">型</span><span class="sxs-lookup"><span data-stu-id="eb503-263">Type</span></span>

*   [<span data-ttu-id="eb503-264">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb503-264">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="eb503-265">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-265">Requirements</span></span>

|<span data-ttu-id="eb503-266">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-266">Requirement</span></span>| <span data-ttu-id="eb503-267">値</span><span class="sxs-lookup"><span data-stu-id="eb503-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-268">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-268">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-269">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-269">1.1</span></span>|
|[<span data-ttu-id="eb503-270">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eb503-270">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="eb503-271">制限あり</span><span class="sxs-lookup"><span data-stu-id="eb503-271">Restricted</span></span>|
|[<span data-ttu-id="eb503-272">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-272">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-273">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="eb503-274">ui: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="eb503-274">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="eb503-275">ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。</span><span class="sxs-lookup"><span data-stu-id="eb503-275">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="eb503-276">型</span><span class="sxs-lookup"><span data-stu-id="eb503-276">Type</span></span>

*   [<span data-ttu-id="eb503-277">UI</span><span class="sxs-lookup"><span data-stu-id="eb503-277">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="eb503-278">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-278">Requirements</span></span>

|<span data-ttu-id="eb503-279">要件</span><span class="sxs-lookup"><span data-stu-id="eb503-279">Requirement</span></span>| <span data-ttu-id="eb503-280">値</span><span class="sxs-lookup"><span data-stu-id="eb503-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb503-281">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eb503-281">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb503-282">1.1</span><span class="sxs-lookup"><span data-stu-id="eb503-282">1.1</span></span>|
|[<span data-ttu-id="eb503-283">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eb503-283">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb503-284">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eb503-284">Compose or Read</span></span>|
