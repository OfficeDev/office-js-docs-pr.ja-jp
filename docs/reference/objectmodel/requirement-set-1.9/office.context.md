---
title: Office コンテキスト要件セット1.9
description: メールボックス API 要件セット1.9 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 6b2657d1e608bd1820d3814d9a6bfab67681824c
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628082"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="f37b4-103">コンテキスト (メールボックス要件セット 1.9)</span><span class="sxs-lookup"><span data-stu-id="f37b4-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="f37b4-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="f37b4-104">[Office](office.md).context</span></span>

<span data-ttu-id="f37b4-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="f37b4-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f37b4-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f37b4-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-107">Requirements</span></span>

|<span data-ttu-id="f37b4-108">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-108">Requirement</span></span>| <span data-ttu-id="f37b4-109">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-111">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-111">1.1</span></span>|
|[<span data-ttu-id="f37b4-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f37b4-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f37b4-114">Properties</span></span>

| <span data-ttu-id="f37b4-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="f37b4-115">Property</span></span> | <span data-ttu-id="f37b4-116">モード</span><span class="sxs-lookup"><span data-stu-id="f37b4-116">Modes</span></span> | <span data-ttu-id="f37b4-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-117">Return type</span></span> | <span data-ttu-id="f37b4-118">最小値</span><span class="sxs-lookup"><span data-stu-id="f37b4-118">Minimum</span></span><br><span data-ttu-id="f37b4-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="f37b4-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f37b4-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="f37b4-120">auth</span></span>](#auth-auth) | <span data-ttu-id="f37b4-121">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-121">Compose</span></span><br><span data-ttu-id="f37b4-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-122">Read</span></span> | [<span data-ttu-id="f37b4-123">Auth</span><span class="sxs-lookup"><span data-stu-id="f37b4-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-124">Identity Api 1.3</span><span class="sxs-lookup"><span data-stu-id="f37b4-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="f37b4-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="f37b4-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="f37b4-126">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-126">Compose</span></span><br><span data-ttu-id="f37b4-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-127">Read</span></span> | <span data-ttu-id="f37b4-128">String</span><span class="sxs-lookup"><span data-stu-id="f37b4-128">String</span></span> | [<span data-ttu-id="f37b4-129">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-130">ダン</span><span class="sxs-lookup"><span data-stu-id="f37b4-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="f37b4-131">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-131">Compose</span></span><br><span data-ttu-id="f37b4-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-132">Read</span></span> | [<span data-ttu-id="f37b4-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f37b4-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="f37b4-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="f37b4-136">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-136">Compose</span></span><br><span data-ttu-id="f37b4-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-137">Read</span></span> | <span data-ttu-id="f37b4-138">String</span><span class="sxs-lookup"><span data-stu-id="f37b4-138">String</span></span> | [<span data-ttu-id="f37b4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-140">主催</span><span class="sxs-lookup"><span data-stu-id="f37b4-140">host</span></span>](#host-hosttype) | <span data-ttu-id="f37b4-141">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-141">Compose</span></span><br><span data-ttu-id="f37b4-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-142">Read</span></span> | [<span data-ttu-id="f37b4-143">HostType</span><span class="sxs-lookup"><span data-stu-id="f37b4-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="f37b4-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="f37b4-146">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-146">Compose</span></span><br><span data-ttu-id="f37b4-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-147">Read</span></span> | [<span data-ttu-id="f37b4-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="f37b4-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-150">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f37b4-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="f37b4-151">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-151">Compose</span></span><br><span data-ttu-id="f37b4-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-152">Read</span></span> | [<span data-ttu-id="f37b4-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f37b4-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-154">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-155">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="f37b4-156">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-156">Compose</span></span><br><span data-ttu-id="f37b4-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-157">Read</span></span> | [<span data-ttu-id="f37b4-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f37b4-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="f37b4-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="f37b4-161">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-161">Compose</span></span><br><span data-ttu-id="f37b4-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-162">Read</span></span> | [<span data-ttu-id="f37b4-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f37b4-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f37b4-165">UI</span><span class="sxs-lookup"><span data-stu-id="f37b4-165">ui</span></span>](#ui-ui) | <span data-ttu-id="f37b4-166">作成</span><span class="sxs-lookup"><span data-stu-id="f37b4-166">Compose</span></span><br><span data-ttu-id="f37b4-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="f37b4-167">Read</span></span> | [<span data-ttu-id="f37b4-168">UI</span><span class="sxs-lookup"><span data-stu-id="f37b4-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="f37b4-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="f37b4-170">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="f37b4-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="f37b4-171">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="f37b4-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="f37b4-172">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office アプリケーションがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="f37b4-173">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="f37b4-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="f37b4-174">「Identity [api 1.3 の要件セット](../../requirement-sets/identity-api-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f37b4-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-175">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-175">Type</span></span>

*   [<span data-ttu-id="f37b4-176">Auth</span><span class="sxs-lookup"><span data-stu-id="f37b4-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="f37b4-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-177">Requirements</span></span>

|<span data-ttu-id="f37b4-178">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-178">Requirement</span></span>| <span data-ttu-id="f37b4-179">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-180">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-181">N/A</span><span class="sxs-lookup"><span data-stu-id="f37b4-181">N/A</span></span>|
|[<span data-ttu-id="f37b4-182">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-183">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-184">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="f37b4-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="f37b4-185">contentLanguage: String</span></span>

<span data-ttu-id="f37b4-186">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="f37b4-187">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-188">型</span><span class="sxs-lookup"><span data-stu-id="f37b4-188">Type</span></span>

*   <span data-ttu-id="f37b4-189">String</span><span class="sxs-lookup"><span data-stu-id="f37b4-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f37b4-190">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-190">Requirements</span></span>

|<span data-ttu-id="f37b4-191">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-191">Requirement</span></span>| <span data-ttu-id="f37b4-192">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-193">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-194">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-194">1.1</span></span>|
|[<span data-ttu-id="f37b4-195">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-196">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-197">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="f37b4-198">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="f37b4-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="f37b4-199">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-200">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-200">Type</span></span>

*   [<span data-ttu-id="f37b4-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="f37b4-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="f37b4-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-202">Requirements</span></span>

|<span data-ttu-id="f37b4-203">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-203">Requirement</span></span>| <span data-ttu-id="f37b4-204">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-205">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-206">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-206">1.1</span></span>|
|[<span data-ttu-id="f37b4-207">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-208">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-209">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-209">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="f37b4-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="f37b4-210">displayLanguage: String</span></span>

<span data-ttu-id="f37b4-211">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="f37b4-212">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-213">型</span><span class="sxs-lookup"><span data-stu-id="f37b4-213">Type</span></span>

*   <span data-ttu-id="f37b4-214">String</span><span class="sxs-lookup"><span data-stu-id="f37b4-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f37b4-215">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-215">Requirements</span></span>

|<span data-ttu-id="f37b4-216">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-216">Requirement</span></span>| <span data-ttu-id="f37b4-217">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-218">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-219">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-219">1.1</span></span>|
|[<span data-ttu-id="f37b4-220">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-221">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-222">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="f37b4-223">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="f37b4-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="f37b4-224">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-224">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-225">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-225">Type</span></span>

*   [<span data-ttu-id="f37b4-226">HostType</span><span class="sxs-lookup"><span data-stu-id="f37b4-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="f37b4-227">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-227">Requirements</span></span>

|<span data-ttu-id="f37b4-228">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-228">Requirement</span></span>| <span data-ttu-id="f37b4-229">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-230">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-231">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-231">1.1</span></span>|
|[<span data-ttu-id="f37b4-232">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-233">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-234">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="f37b4-235">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="f37b4-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="f37b4-236">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-236">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-237">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-237">Type</span></span>

*   [<span data-ttu-id="f37b4-238">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f37b4-238">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="f37b4-239">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-239">Requirements</span></span>

|<span data-ttu-id="f37b4-240">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-240">Requirement</span></span>| <span data-ttu-id="f37b4-241">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-242">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-243">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-243">1.1</span></span>|
|[<span data-ttu-id="f37b4-244">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-245">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-245">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-246">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-246">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="f37b4-247">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="f37b4-247">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="f37b4-248">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-248">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-249">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-249">Type</span></span>

*   [<span data-ttu-id="f37b4-250">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="f37b4-250">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="f37b4-251">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-251">Requirements</span></span>

|<span data-ttu-id="f37b4-252">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-252">Requirement</span></span>| <span data-ttu-id="f37b4-253">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-254">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-254">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-255">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-255">1.1</span></span>|
|[<span data-ttu-id="f37b4-256">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-256">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-257">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-257">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f37b4-258">例</span><span class="sxs-lookup"><span data-stu-id="f37b4-258">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="f37b4-259">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="f37b4-259">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="f37b4-260">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-260">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="f37b4-261">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="f37b4-261">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-262">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-262">Type</span></span>

*   [<span data-ttu-id="f37b4-263">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f37b4-263">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="f37b4-264">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-264">Requirements</span></span>

|<span data-ttu-id="f37b4-265">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-265">Requirement</span></span>| <span data-ttu-id="f37b4-266">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-267">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-267">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-268">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-268">1.1</span></span>|
|[<span data-ttu-id="f37b4-269">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="f37b4-269">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="f37b4-270">制限あり</span><span class="sxs-lookup"><span data-stu-id="f37b4-270">Restricted</span></span>|
|[<span data-ttu-id="f37b4-271">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-271">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-272">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-272">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="f37b4-273">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="f37b4-273">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="f37b4-274">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="f37b4-274">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="f37b4-275">種類</span><span class="sxs-lookup"><span data-stu-id="f37b4-275">Type</span></span>

*   [<span data-ttu-id="f37b4-276">UI</span><span class="sxs-lookup"><span data-stu-id="f37b4-276">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="f37b4-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="f37b4-277">Requirements</span></span>

|<span data-ttu-id="f37b4-278">要件</span><span class="sxs-lookup"><span data-stu-id="f37b4-278">Requirement</span></span>| <span data-ttu-id="f37b4-279">値</span><span class="sxs-lookup"><span data-stu-id="f37b4-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="f37b4-280">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="f37b4-280">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f37b4-281">1.1</span><span class="sxs-lookup"><span data-stu-id="f37b4-281">1.1</span></span>|
|[<span data-ttu-id="f37b4-282">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="f37b4-282">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f37b4-283">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="f37b4-283">Compose or Read</span></span>|
