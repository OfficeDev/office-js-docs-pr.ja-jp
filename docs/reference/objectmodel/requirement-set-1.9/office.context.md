---
title: Office コンテキスト要件セット1.9
description: メールボックス API 要件セット1.9 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 3a8a9fe65ebf3c5a5ee63766f71dfce8e3f8d905
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570724"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="dce37-103">コンテキスト (メールボックス要件セット 1.9)</span><span class="sxs-lookup"><span data-stu-id="dce37-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="dce37-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="dce37-104">[Office](office.md).context</span></span>

<span data-ttu-id="dce37-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="dce37-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="dce37-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dce37-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dce37-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-107">Requirements</span></span>

|<span data-ttu-id="dce37-108">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-108">Requirement</span></span>| <span data-ttu-id="dce37-109">値</span><span class="sxs-lookup"><span data-stu-id="dce37-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-111">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-111">1.1</span></span>|
|[<span data-ttu-id="dce37-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dce37-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dce37-114">Properties</span></span>

| <span data-ttu-id="dce37-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="dce37-115">Property</span></span> | <span data-ttu-id="dce37-116">モード</span><span class="sxs-lookup"><span data-stu-id="dce37-116">Modes</span></span> | <span data-ttu-id="dce37-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="dce37-117">Return type</span></span> | <span data-ttu-id="dce37-118">最小値</span><span class="sxs-lookup"><span data-stu-id="dce37-118">Minimum</span></span><br><span data-ttu-id="dce37-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="dce37-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dce37-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="dce37-120">auth</span></span>](#auth-auth) | <span data-ttu-id="dce37-121">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-121">Compose</span></span><br><span data-ttu-id="dce37-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-122">Read</span></span> | [<span data-ttu-id="dce37-123">Auth</span><span class="sxs-lookup"><span data-stu-id="dce37-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-124">Identity Api 1.3</span><span class="sxs-lookup"><span data-stu-id="dce37-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="dce37-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="dce37-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="dce37-126">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-126">Compose</span></span><br><span data-ttu-id="dce37-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-127">Read</span></span> | <span data-ttu-id="dce37-128">文字列</span><span class="sxs-lookup"><span data-stu-id="dce37-128">String</span></span> | [<span data-ttu-id="dce37-129">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-130">ダン</span><span class="sxs-lookup"><span data-stu-id="dce37-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="dce37-131">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-131">Compose</span></span><br><span data-ttu-id="dce37-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-132">Read</span></span> | [<span data-ttu-id="dce37-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dce37-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="dce37-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="dce37-136">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-136">Compose</span></span><br><span data-ttu-id="dce37-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-137">Read</span></span> | <span data-ttu-id="dce37-138">文字列</span><span class="sxs-lookup"><span data-stu-id="dce37-138">String</span></span> | [<span data-ttu-id="dce37-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-140">主催</span><span class="sxs-lookup"><span data-stu-id="dce37-140">host</span></span>](#host-hosttype) | <span data-ttu-id="dce37-141">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-141">Compose</span></span><br><span data-ttu-id="dce37-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-142">Read</span></span> | [<span data-ttu-id="dce37-143">HostType</span><span class="sxs-lookup"><span data-stu-id="dce37-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-144">1.5</span><span class="sxs-lookup"><span data-stu-id="dce37-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dce37-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="dce37-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="dce37-146">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-146">Compose</span></span><br><span data-ttu-id="dce37-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-147">Read</span></span> | [<span data-ttu-id="dce37-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="dce37-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-150">platform</span><span class="sxs-lookup"><span data-stu-id="dce37-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="dce37-151">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-151">Compose</span></span><br><span data-ttu-id="dce37-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-152">Read</span></span> | [<span data-ttu-id="dce37-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dce37-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-154">1.5</span><span class="sxs-lookup"><span data-stu-id="dce37-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dce37-155">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="dce37-156">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-156">Compose</span></span><br><span data-ttu-id="dce37-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-157">Read</span></span> | [<span data-ttu-id="dce37-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dce37-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-159">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="dce37-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="dce37-161">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-161">Compose</span></span><br><span data-ttu-id="dce37-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-162">Read</span></span> | [<span data-ttu-id="dce37-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dce37-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-164">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dce37-165">UI</span><span class="sxs-lookup"><span data-stu-id="dce37-165">ui</span></span>](#ui-ui) | <span data-ttu-id="dce37-166">作成</span><span class="sxs-lookup"><span data-stu-id="dce37-166">Compose</span></span><br><span data-ttu-id="dce37-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="dce37-167">Read</span></span> | [<span data-ttu-id="dce37-168">UI</span><span class="sxs-lookup"><span data-stu-id="dce37-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="dce37-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="dce37-170">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="dce37-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="dce37-171">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="dce37-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="dce37-172">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office アプリケーションがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="dce37-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="dce37-173">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="dce37-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="dce37-174">「Identity [api 1.3 の要件セット](../../requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dce37-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-175">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-175">Type</span></span>

*   [<span data-ttu-id="dce37-176">Auth</span><span class="sxs-lookup"><span data-stu-id="dce37-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="dce37-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-177">Requirements</span></span>

|<span data-ttu-id="dce37-178">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-178">Requirement</span></span>| <span data-ttu-id="dce37-179">値</span><span class="sxs-lookup"><span data-stu-id="dce37-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-181">N/A</span><span class="sxs-lookup"><span data-stu-id="dce37-181">N/A</span></span>|
|[<span data-ttu-id="dce37-182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-183">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-184">例</span><span class="sxs-lookup"><span data-stu-id="dce37-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="dce37-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="dce37-185">contentLanguage: String</span></span>

<span data-ttu-id="dce37-186">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="dce37-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="dce37-187">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="dce37-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-188">型</span><span class="sxs-lookup"><span data-stu-id="dce37-188">Type</span></span>

*   <span data-ttu-id="dce37-189">String</span><span class="sxs-lookup"><span data-stu-id="dce37-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dce37-190">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-190">Requirements</span></span>

|<span data-ttu-id="dce37-191">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-191">Requirement</span></span>| <span data-ttu-id="dce37-192">値</span><span class="sxs-lookup"><span data-stu-id="dce37-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-194">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-194">1.1</span></span>|
|[<span data-ttu-id="dce37-195">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-196">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-197">例</span><span class="sxs-lookup"><span data-stu-id="dce37-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="dce37-198">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="dce37-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="dce37-199">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="dce37-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-200">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-200">Type</span></span>

*   [<span data-ttu-id="dce37-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dce37-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="dce37-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-202">Requirements</span></span>

|<span data-ttu-id="dce37-203">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-203">Requirement</span></span>| <span data-ttu-id="dce37-204">値</span><span class="sxs-lookup"><span data-stu-id="dce37-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-206">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-206">1.1</span></span>|
|[<span data-ttu-id="dce37-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-209">例</span><span class="sxs-lookup"><span data-stu-id="dce37-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="dce37-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="dce37-210">displayLanguage: String</span></span>

<span data-ttu-id="dce37-211">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="dce37-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="dce37-212">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="dce37-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-213">型</span><span class="sxs-lookup"><span data-stu-id="dce37-213">Type</span></span>

*   <span data-ttu-id="dce37-214">String</span><span class="sxs-lookup"><span data-stu-id="dce37-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dce37-215">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-215">Requirements</span></span>

|<span data-ttu-id="dce37-216">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-216">Requirement</span></span>| <span data-ttu-id="dce37-217">値</span><span class="sxs-lookup"><span data-stu-id="dce37-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-219">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-219">1.1</span></span>|
|[<span data-ttu-id="dce37-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-221">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-222">例</span><span class="sxs-lookup"><span data-stu-id="dce37-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="dce37-223">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="dce37-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="dce37-224">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="dce37-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="dce37-225">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="dce37-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-226">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-226">Type</span></span>

*   [<span data-ttu-id="dce37-227">HostType</span><span class="sxs-lookup"><span data-stu-id="dce37-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="dce37-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-228">Requirements</span></span>

|<span data-ttu-id="dce37-229">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-229">Requirement</span></span>| <span data-ttu-id="dce37-230">値</span><span class="sxs-lookup"><span data-stu-id="dce37-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-231">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-232">1.5</span><span class="sxs-lookup"><span data-stu-id="dce37-232">1.5</span></span>|
|[<span data-ttu-id="dce37-233">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-234">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-235">例</span><span class="sxs-lookup"><span data-stu-id="dce37-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="dce37-236">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="dce37-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="dce37-237">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="dce37-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="dce37-238">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="dce37-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-239">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-239">Type</span></span>

*   [<span data-ttu-id="dce37-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="dce37-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="dce37-241">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-241">Requirements</span></span>

|<span data-ttu-id="dce37-242">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-242">Requirement</span></span>| <span data-ttu-id="dce37-243">値</span><span class="sxs-lookup"><span data-stu-id="dce37-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-244">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-245">1.5</span><span class="sxs-lookup"><span data-stu-id="dce37-245">1.5</span></span>|
|[<span data-ttu-id="dce37-246">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-247">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-248">例</span><span class="sxs-lookup"><span data-stu-id="dce37-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="dce37-249">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="dce37-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="dce37-250">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="dce37-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-251">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-251">Type</span></span>

*   [<span data-ttu-id="dce37-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dce37-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="dce37-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-253">Requirements</span></span>

|<span data-ttu-id="dce37-254">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-254">Requirement</span></span>| <span data-ttu-id="dce37-255">値</span><span class="sxs-lookup"><span data-stu-id="dce37-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-256">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-257">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-257">1.1</span></span>|
|[<span data-ttu-id="dce37-258">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-259">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dce37-260">例</span><span class="sxs-lookup"><span data-stu-id="dce37-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="dce37-261">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="dce37-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="dce37-262">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="dce37-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="dce37-263">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="dce37-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-264">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-264">Type</span></span>

*   [<span data-ttu-id="dce37-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dce37-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="dce37-266">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-266">Requirements</span></span>

|<span data-ttu-id="dce37-267">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-267">Requirement</span></span>| <span data-ttu-id="dce37-268">値</span><span class="sxs-lookup"><span data-stu-id="dce37-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-269">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-270">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-270">1.1</span></span>|
|[<span data-ttu-id="dce37-271">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="dce37-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="dce37-272">制限あり</span><span class="sxs-lookup"><span data-stu-id="dce37-272">Restricted</span></span>|
|[<span data-ttu-id="dce37-273">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-274">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="dce37-275">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="dce37-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="dce37-276">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="dce37-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="dce37-277">種類</span><span class="sxs-lookup"><span data-stu-id="dce37-277">Type</span></span>

*   [<span data-ttu-id="dce37-278">UI</span><span class="sxs-lookup"><span data-stu-id="dce37-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="dce37-279">Requirements</span><span class="sxs-lookup"><span data-stu-id="dce37-279">Requirements</span></span>

|<span data-ttu-id="dce37-280">要件</span><span class="sxs-lookup"><span data-stu-id="dce37-280">Requirement</span></span>| <span data-ttu-id="dce37-281">値</span><span class="sxs-lookup"><span data-stu-id="dce37-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="dce37-282">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="dce37-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dce37-283">1.1</span><span class="sxs-lookup"><span data-stu-id="dce37-283">1.1</span></span>|
|[<span data-ttu-id="dce37-284">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="dce37-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dce37-285">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="dce37-285">Compose or Read</span></span>|
