---
title: Office コンテキスト-プレビュー要件セット
description: メールボックス API プレビュー要件セットを使用して Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5987f81b0b4790b74bde092fc3de44df4fa3ed16
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293815"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="4a67f-103">コンテキスト (メールボックスプレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="4a67f-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="4a67f-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="4a67f-104">[Office](office.md).context</span></span>

<span data-ttu-id="4a67f-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="4a67f-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4a67f-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a67f-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-107">Requirements</span></span>

|<span data-ttu-id="4a67f-108">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-108">Requirement</span></span>| <span data-ttu-id="4a67f-109">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-111">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-111">1.1</span></span>|
|[<span data-ttu-id="4a67f-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4a67f-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="4a67f-114">Properties</span></span>

| <span data-ttu-id="4a67f-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="4a67f-115">Property</span></span> | <span data-ttu-id="4a67f-116">モード</span><span class="sxs-lookup"><span data-stu-id="4a67f-116">Modes</span></span> | <span data-ttu-id="4a67f-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="4a67f-117">Return type</span></span> | <span data-ttu-id="4a67f-118">最小値</span><span class="sxs-lookup"><span data-stu-id="4a67f-118">Minimum</span></span><br><span data-ttu-id="4a67f-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="4a67f-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4a67f-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="4a67f-120">auth</span></span>](#auth-auth) | <span data-ttu-id="4a67f-121">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-121">Compose</span></span><br><span data-ttu-id="4a67f-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-122">Read</span></span> | [<span data-ttu-id="4a67f-123">Auth</span><span class="sxs-lookup"><span data-stu-id="4a67f-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="4a67f-124">プレビュー</span><span class="sxs-lookup"><span data-stu-id="4a67f-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="4a67f-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="4a67f-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="4a67f-126">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-126">Compose</span></span><br><span data-ttu-id="4a67f-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-127">Read</span></span> | <span data-ttu-id="4a67f-128">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-128">String</span></span> | [<span data-ttu-id="4a67f-129">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-130">ダン</span><span class="sxs-lookup"><span data-stu-id="4a67f-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="4a67f-131">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-131">Compose</span></span><br><span data-ttu-id="4a67f-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-132">Read</span></span> | [<span data-ttu-id="4a67f-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="4a67f-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="4a67f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="4a67f-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="4a67f-136">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-136">Compose</span></span><br><span data-ttu-id="4a67f-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-137">Read</span></span> | <span data-ttu-id="4a67f-138">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-138">String</span></span> | [<span data-ttu-id="4a67f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-140">主催</span><span class="sxs-lookup"><span data-stu-id="4a67f-140">host</span></span>](#host-hosttype) | <span data-ttu-id="4a67f-141">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-141">Compose</span></span><br><span data-ttu-id="4a67f-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-142">Read</span></span> | [<span data-ttu-id="4a67f-143">HostType</span><span class="sxs-lookup"><span data-stu-id="4a67f-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="4a67f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="4a67f-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="4a67f-146">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-146">Compose</span></span><br><span data-ttu-id="4a67f-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-147">Read</span></span> | [<span data-ttu-id="4a67f-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="4a67f-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="4a67f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="4a67f-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="4a67f-151">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-151">Compose</span></span><br><span data-ttu-id="4a67f-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-152">Read</span></span> | [<span data-ttu-id="4a67f-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="4a67f-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="4a67f-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="4a67f-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="4a67f-155">platform</span><span class="sxs-lookup"><span data-stu-id="4a67f-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="4a67f-156">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-156">Compose</span></span><br><span data-ttu-id="4a67f-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-157">Read</span></span> | [<span data-ttu-id="4a67f-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="4a67f-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="4a67f-159">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-160">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="4a67f-161">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-161">Compose</span></span><br><span data-ttu-id="4a67f-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-162">Read</span></span> | [<span data-ttu-id="4a67f-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="4a67f-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="4a67f-164">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a67f-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="4a67f-166">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-166">Compose</span></span><br><span data-ttu-id="4a67f-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-167">Read</span></span> | [<span data-ttu-id="4a67f-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a67f-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="4a67f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4a67f-170">UI</span><span class="sxs-lookup"><span data-stu-id="4a67f-170">ui</span></span>](#ui-ui) | <span data-ttu-id="4a67f-171">作成</span><span class="sxs-lookup"><span data-stu-id="4a67f-171">Compose</span></span><br><span data-ttu-id="4a67f-172">読み取り</span><span class="sxs-lookup"><span data-stu-id="4a67f-172">Read</span></span> | [<span data-ttu-id="4a67f-173">UI</span><span class="sxs-lookup"><span data-stu-id="4a67f-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="4a67f-174">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="4a67f-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="4a67f-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="4a67f-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="4a67f-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="4a67f-177">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office アプリケーションがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="4a67f-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="4a67f-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-179">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-179">Type</span></span>

*   [<span data-ttu-id="4a67f-180">Auth</span><span class="sxs-lookup"><span data-stu-id="4a67f-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="4a67f-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-181">Requirements</span></span>

|<span data-ttu-id="4a67f-182">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-182">Requirement</span></span>| <span data-ttu-id="4a67f-183">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="4a67f-185">Preview</span></span>|
|[<span data-ttu-id="4a67f-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-188">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="4a67f-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="4a67f-189">contentLanguage: String</span></span>

<span data-ttu-id="4a67f-190">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="4a67f-191">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-192">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-192">Type</span></span>

*   <span data-ttu-id="4a67f-193">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a67f-194">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-194">Requirements</span></span>

|<span data-ttu-id="4a67f-195">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-195">Requirement</span></span>| <span data-ttu-id="4a67f-196">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-198">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-198">1.1</span></span>|
|[<span data-ttu-id="4a67f-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-201">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="4a67f-202">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="4a67f-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="4a67f-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-204">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-204">Type</span></span>

*   [<span data-ttu-id="4a67f-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="4a67f-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="4a67f-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-206">Requirements</span></span>

|<span data-ttu-id="4a67f-207">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-207">Requirement</span></span>| <span data-ttu-id="4a67f-208">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-210">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-210">1.1</span></span>|
|[<span data-ttu-id="4a67f-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-213">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="4a67f-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="4a67f-214">displayLanguage: String</span></span>

<span data-ttu-id="4a67f-215">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="4a67f-216">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-217">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-217">Type</span></span>

*   <span data-ttu-id="4a67f-218">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a67f-219">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-219">Requirements</span></span>

|<span data-ttu-id="4a67f-220">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-220">Requirement</span></span>| <span data-ttu-id="4a67f-221">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-223">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-223">1.1</span></span>|
|[<span data-ttu-id="4a67f-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-226">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="4a67f-227">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="4a67f-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="4a67f-228">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-229">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-229">Type</span></span>

*   [<span data-ttu-id="4a67f-230">HostType</span><span class="sxs-lookup"><span data-stu-id="4a67f-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="4a67f-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-231">Requirements</span></span>

|<span data-ttu-id="4a67f-232">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-232">Requirement</span></span>| <span data-ttu-id="4a67f-233">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-235">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-235">1.1</span></span>|
|[<span data-ttu-id="4a67f-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-238">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="4a67f-239">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="4a67f-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="4a67f-240">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="4a67f-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="4a67f-241">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="4a67f-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="4a67f-242">Office テーマの色を使用すると、アドインの配色を、[ **ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office クライアントアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="4a67f-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="4a67f-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="4a67f-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-244">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-244">Type</span></span>

*   [<span data-ttu-id="4a67f-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="4a67f-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="4a67f-246">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="4a67f-246">Properties:</span></span>

|<span data-ttu-id="4a67f-247">名前</span><span class="sxs-lookup"><span data-stu-id="4a67f-247">Name</span></span>| <span data-ttu-id="4a67f-248">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-248">Type</span></span>| <span data-ttu-id="4a67f-249">説明</span><span class="sxs-lookup"><span data-stu-id="4a67f-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="4a67f-250">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-250">String</span></span>|<span data-ttu-id="4a67f-251">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="4a67f-252">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-252">String</span></span>|<span data-ttu-id="4a67f-253">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="4a67f-254">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-254">String</span></span>|<span data-ttu-id="4a67f-255">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="4a67f-256">String</span><span class="sxs-lookup"><span data-stu-id="4a67f-256">String</span></span>|<span data-ttu-id="4a67f-257">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a67f-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-258">Requirements</span></span>

|<span data-ttu-id="4a67f-259">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-259">Requirement</span></span>| <span data-ttu-id="4a67f-260">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-262">プレビュー</span><span class="sxs-lookup"><span data-stu-id="4a67f-262">Preview</span></span>|
|[<span data-ttu-id="4a67f-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-265">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="4a67f-266">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="4a67f-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="4a67f-267">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-268">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-268">Type</span></span>

*   [<span data-ttu-id="4a67f-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="4a67f-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="4a67f-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-270">Requirements</span></span>

|<span data-ttu-id="4a67f-271">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-271">Requirement</span></span>| <span data-ttu-id="4a67f-272">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-274">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-274">1.1</span></span>|
|[<span data-ttu-id="4a67f-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-277">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="4a67f-278">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="4a67f-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="4a67f-279">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-280">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-280">Type</span></span>

*   [<span data-ttu-id="4a67f-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="4a67f-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="4a67f-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-282">Requirements</span></span>

|<span data-ttu-id="4a67f-283">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-283">Requirement</span></span>| <span data-ttu-id="4a67f-284">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-286">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-286">1.1</span></span>|
|[<span data-ttu-id="4a67f-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-288">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a67f-289">例</span><span class="sxs-lookup"><span data-stu-id="4a67f-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="4a67f-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="4a67f-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="4a67f-291">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="4a67f-292">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="4a67f-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-293">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-293">Type</span></span>

*   [<span data-ttu-id="4a67f-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4a67f-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="4a67f-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-295">Requirements</span></span>

|<span data-ttu-id="4a67f-296">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-296">Requirement</span></span>| <span data-ttu-id="4a67f-297">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-299">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-299">1.1</span></span>|
|[<span data-ttu-id="4a67f-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="4a67f-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="4a67f-301">制限あり</span><span class="sxs-lookup"><span data-stu-id="4a67f-301">Restricted</span></span>|
|[<span data-ttu-id="4a67f-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-303">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="4a67f-304">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="4a67f-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="4a67f-305">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="4a67f-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="4a67f-306">型</span><span class="sxs-lookup"><span data-stu-id="4a67f-306">Type</span></span>

*   [<span data-ttu-id="4a67f-307">UI</span><span class="sxs-lookup"><span data-stu-id="4a67f-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="4a67f-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a67f-308">Requirements</span></span>

|<span data-ttu-id="4a67f-309">要件</span><span class="sxs-lookup"><span data-stu-id="4a67f-309">Requirement</span></span>| <span data-ttu-id="4a67f-310">値</span><span class="sxs-lookup"><span data-stu-id="4a67f-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a67f-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="4a67f-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4a67f-312">1.1</span><span class="sxs-lookup"><span data-stu-id="4a67f-312">1.1</span></span>|
|[<span data-ttu-id="4a67f-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="4a67f-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4a67f-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="4a67f-314">Compose or Read</span></span>|
