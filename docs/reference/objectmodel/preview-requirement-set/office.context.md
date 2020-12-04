---
title: Office コンテキスト-プレビュー要件セット
description: メールボックス API プレビュー要件セットを使用して Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8370df907aa3ab0534254057860c187cec583e6c
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570787"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="ee5b1-103">コンテキスト (メールボックスプレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ee5b1-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ee5b1-104">[Office](office.md).context</span></span>

<span data-ttu-id="ee5b1-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ee5b1-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5b1-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-107">Requirements</span></span>

|<span data-ttu-id="ee5b1-108">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-108">Requirement</span></span>| <span data-ttu-id="ee5b1-109">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-111">1.1</span></span>|
|[<span data-ttu-id="ee5b1-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ee5b1-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ee5b1-114">Properties</span></span>

| <span data-ttu-id="ee5b1-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ee5b1-115">Property</span></span> | <span data-ttu-id="ee5b1-116">モード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-116">Modes</span></span> | <span data-ttu-id="ee5b1-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-117">Return type</span></span> | <span data-ttu-id="ee5b1-118">最小値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-118">Minimum</span></span><br><span data-ttu-id="ee5b1-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="ee5b1-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ee5b1-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="ee5b1-120">auth</span></span>](#auth-auth) | <span data-ttu-id="ee5b1-121">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-121">Compose</span></span><br><span data-ttu-id="ee5b1-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-122">Read</span></span> | [<span data-ttu-id="ee5b1-123">Auth</span><span class="sxs-lookup"><span data-stu-id="ee5b1-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-124">Identity Api 1.3</span><span class="sxs-lookup"><span data-stu-id="ee5b1-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="ee5b1-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ee5b1-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ee5b1-126">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-126">Compose</span></span><br><span data-ttu-id="ee5b1-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-127">Read</span></span> | <span data-ttu-id="ee5b1-128">文字列</span><span class="sxs-lookup"><span data-stu-id="ee5b1-128">String</span></span> | [<span data-ttu-id="ee5b1-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-130">ダン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ee5b1-131">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-131">Compose</span></span><br><span data-ttu-id="ee5b1-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-132">Read</span></span> | [<span data-ttu-id="ee5b1-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ee5b1-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ee5b1-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ee5b1-136">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-136">Compose</span></span><br><span data-ttu-id="ee5b1-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-137">Read</span></span> | <span data-ttu-id="ee5b1-138">文字列</span><span class="sxs-lookup"><span data-stu-id="ee5b1-138">String</span></span> | [<span data-ttu-id="ee5b1-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-140">主催</span><span class="sxs-lookup"><span data-stu-id="ee5b1-140">host</span></span>](#host-hosttype) | <span data-ttu-id="ee5b1-141">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-141">Compose</span></span><br><span data-ttu-id="ee5b1-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-142">Read</span></span> | [<span data-ttu-id="ee5b1-143">HostType</span><span class="sxs-lookup"><span data-stu-id="ee5b1-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-144">1.5</span><span class="sxs-lookup"><span data-stu-id="ee5b1-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ee5b1-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="ee5b1-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ee5b1-146">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-146">Compose</span></span><br><span data-ttu-id="ee5b1-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-147">Read</span></span> | [<span data-ttu-id="ee5b1-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="ee5b1-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="ee5b1-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="ee5b1-151">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-151">Compose</span></span><br><span data-ttu-id="ee5b1-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-152">Read</span></span> | [<span data-ttu-id="ee5b1-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ee5b1-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ee5b1-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="ee5b1-155">platform</span><span class="sxs-lookup"><span data-stu-id="ee5b1-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ee5b1-156">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-156">Compose</span></span><br><span data-ttu-id="ee5b1-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-157">Read</span></span> | [<span data-ttu-id="ee5b1-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ee5b1-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-159">1.5</span><span class="sxs-lookup"><span data-stu-id="ee5b1-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ee5b1-160">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ee5b1-161">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-161">Compose</span></span><br><span data-ttu-id="ee5b1-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-162">Read</span></span> | [<span data-ttu-id="ee5b1-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ee5b1-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ee5b1-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ee5b1-166">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-166">Compose</span></span><br><span data-ttu-id="ee5b1-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-167">Read</span></span> | [<span data-ttu-id="ee5b1-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ee5b1-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-169">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ee5b1-170">UI</span><span class="sxs-lookup"><span data-stu-id="ee5b1-170">ui</span></span>](#ui-ui) | <span data-ttu-id="ee5b1-171">作成</span><span class="sxs-lookup"><span data-stu-id="ee5b1-171">Compose</span></span><br><span data-ttu-id="ee5b1-172">読み取り</span><span class="sxs-lookup"><span data-stu-id="ee5b1-172">Read</span></span> | [<span data-ttu-id="ee5b1-173">UI</span><span class="sxs-lookup"><span data-stu-id="ee5b1-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ee5b1-174">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ee5b1-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="ee5b1-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="ee5b1-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="ee5b1-177">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office アプリケーションがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="ee5b1-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-179">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-179">Type</span></span>

*   [<span data-ttu-id="ee5b1-180">Auth</span><span class="sxs-lookup"><span data-stu-id="ee5b1-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-181">Requirements</span></span>

|<span data-ttu-id="ee5b1-182">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-182">Requirement</span></span>| <span data-ttu-id="ee5b1-183">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ee5b1-185">Preview</span></span>|
|[<span data-ttu-id="ee5b1-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-188">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="ee5b1-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-189">contentLanguage: String</span></span>

<span data-ttu-id="ee5b1-190">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ee5b1-191">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-192">型</span><span class="sxs-lookup"><span data-stu-id="ee5b1-192">Type</span></span>

*   <span data-ttu-id="ee5b1-193">String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5b1-194">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-194">Requirements</span></span>

|<span data-ttu-id="ee5b1-195">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-195">Requirement</span></span>| <span data-ttu-id="ee5b1-196">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-198">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-198">1.1</span></span>|
|[<span data-ttu-id="ee5b1-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-201">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ee5b1-202">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ee5b1-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-204">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-204">Type</span></span>

*   [<span data-ttu-id="ee5b1-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ee5b1-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-206">Requirements</span></span>

|<span data-ttu-id="ee5b1-207">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-207">Requirement</span></span>| <span data-ttu-id="ee5b1-208">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-210">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-210">1.1</span></span>|
|[<span data-ttu-id="ee5b1-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-213">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ee5b1-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-214">displayLanguage: String</span></span>

<span data-ttu-id="ee5b1-215">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ee5b1-216">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-217">型</span><span class="sxs-lookup"><span data-stu-id="ee5b1-217">Type</span></span>

*   <span data-ttu-id="ee5b1-218">String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5b1-219">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-219">Requirements</span></span>

|<span data-ttu-id="ee5b1-220">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-220">Requirement</span></span>| <span data-ttu-id="ee5b1-221">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-223">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-223">1.1</span></span>|
|[<span data-ttu-id="ee5b1-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-226">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ee5b1-227">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ee5b1-228">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="ee5b1-229">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してホストを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-230">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-230">Type</span></span>

*   [<span data-ttu-id="ee5b1-231">HostType</span><span class="sxs-lookup"><span data-stu-id="ee5b1-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-232">Requirements</span></span>

|<span data-ttu-id="ee5b1-233">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-233">Requirement</span></span>| <span data-ttu-id="ee5b1-234">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-236">1.5</span><span class="sxs-lookup"><span data-stu-id="ee5b1-236">1.5</span></span>|
|[<span data-ttu-id="ee5b1-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-239">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="ee5b1-240">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="ee5b1-241">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="ee5b1-242">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="ee5b1-243">Office テーマの色を使用すると、アドインの配色を、[ **ファイル > Office アカウント > Office テーマ UI** を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office クライアントアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="ee5b1-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="ee5b1-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-245">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-245">Type</span></span>

*   [<span data-ttu-id="ee5b1-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ee5b1-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="ee5b1-247">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="ee5b1-247">Properties:</span></span>

|<span data-ttu-id="ee5b1-248">名前</span><span class="sxs-lookup"><span data-stu-id="ee5b1-248">Name</span></span>| <span data-ttu-id="ee5b1-249">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-249">Type</span></span>| <span data-ttu-id="ee5b1-250">説明</span><span class="sxs-lookup"><span data-stu-id="ee5b1-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="ee5b1-251">文字列</span><span class="sxs-lookup"><span data-stu-id="ee5b1-251">String</span></span>|<span data-ttu-id="ee5b1-252">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="ee5b1-253">String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-253">String</span></span>|<span data-ttu-id="ee5b1-254">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="ee5b1-255">String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-255">String</span></span>|<span data-ttu-id="ee5b1-256">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="ee5b1-257">String</span><span class="sxs-lookup"><span data-stu-id="ee5b1-257">String</span></span>|<span data-ttu-id="ee5b1-258">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ee5b1-259">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-259">Requirements</span></span>

|<span data-ttu-id="ee5b1-260">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-260">Requirement</span></span>| <span data-ttu-id="ee5b1-261">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-262">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-263">プレビュー</span><span class="sxs-lookup"><span data-stu-id="ee5b1-263">Preview</span></span>|
|[<span data-ttu-id="ee5b1-264">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-265">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-266">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="ee5b1-267">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ee5b1-268">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="ee5b1-269">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-270">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-270">Type</span></span>

*   [<span data-ttu-id="ee5b1-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ee5b1-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-272">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-272">Requirements</span></span>

|<span data-ttu-id="ee5b1-273">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-273">Requirement</span></span>| <span data-ttu-id="ee5b1-274">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-275">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-276">1.5</span><span class="sxs-lookup"><span data-stu-id="ee5b1-276">1.5</span></span>|
|[<span data-ttu-id="ee5b1-277">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-278">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-279">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ee5b1-280">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ee5b1-281">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-282">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-282">Type</span></span>

*   [<span data-ttu-id="ee5b1-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ee5b1-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-284">Requirements</span></span>

|<span data-ttu-id="ee5b1-285">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-285">Requirement</span></span>| <span data-ttu-id="ee5b1-286">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-287">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-288">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-288">1.1</span></span>|
|[<span data-ttu-id="ee5b1-289">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-290">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5b1-291">例</span><span class="sxs-lookup"><span data-stu-id="ee5b1-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ee5b1-292">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ee5b1-293">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ee5b1-294">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-295">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-295">Type</span></span>

*   [<span data-ttu-id="ee5b1-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ee5b1-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-297">Requirements</span></span>

|<span data-ttu-id="ee5b1-298">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-298">Requirement</span></span>| <span data-ttu-id="ee5b1-299">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-300">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-301">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-301">1.1</span></span>|
|[<span data-ttu-id="ee5b1-302">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ee5b1-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ee5b1-303">制限あり</span><span class="sxs-lookup"><span data-stu-id="ee5b1-303">Restricted</span></span>|
|[<span data-ttu-id="ee5b1-304">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-305">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ee5b1-306">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ee5b1-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ee5b1-307">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="ee5b1-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5b1-308">種類</span><span class="sxs-lookup"><span data-stu-id="ee5b1-308">Type</span></span>

*   [<span data-ttu-id="ee5b1-309">UI</span><span class="sxs-lookup"><span data-stu-id="ee5b1-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ee5b1-310">Requirements</span><span class="sxs-lookup"><span data-stu-id="ee5b1-310">Requirements</span></span>

|<span data-ttu-id="ee5b1-311">要件</span><span class="sxs-lookup"><span data-stu-id="ee5b1-311">Requirement</span></span>| <span data-ttu-id="ee5b1-312">値</span><span class="sxs-lookup"><span data-stu-id="ee5b1-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5b1-313">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ee5b1-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ee5b1-314">1.1</span><span class="sxs-lookup"><span data-stu-id="ee5b1-314">1.1</span></span>|
|[<span data-ttu-id="ee5b1-315">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ee5b1-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ee5b1-316">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ee5b1-316">Compose or Read</span></span>|
