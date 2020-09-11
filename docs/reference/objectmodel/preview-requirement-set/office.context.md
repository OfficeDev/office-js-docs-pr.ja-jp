---
title: Office コンテキスト-プレビュー要件セット
description: メールボックス API プレビュー要件セットを使用して Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 64a96336ec181747fecf06c8cd2441b600ac8a10
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431116"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="d72d0-103">コンテキスト (メールボックスプレビュー要件セット)</span><span class="sxs-lookup"><span data-stu-id="d72d0-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="d72d0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="d72d0-104">[Office](office.md).context</span></span>

<span data-ttu-id="d72d0-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="d72d0-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d72d0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d72d0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-107">Requirements</span></span>

|<span data-ttu-id="d72d0-108">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-108">Requirement</span></span>| <span data-ttu-id="d72d0-109">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-111">1.1</span></span>|
|[<span data-ttu-id="d72d0-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d72d0-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="d72d0-114">Properties</span></span>

| <span data-ttu-id="d72d0-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="d72d0-115">Property</span></span> | <span data-ttu-id="d72d0-116">モード</span><span class="sxs-lookup"><span data-stu-id="d72d0-116">Modes</span></span> | <span data-ttu-id="d72d0-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-117">Return type</span></span> | <span data-ttu-id="d72d0-118">最小値</span><span class="sxs-lookup"><span data-stu-id="d72d0-118">Minimum</span></span><br><span data-ttu-id="d72d0-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="d72d0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d72d0-120">authoritative</span><span class="sxs-lookup"><span data-stu-id="d72d0-120">auth</span></span>](#auth-auth) | <span data-ttu-id="d72d0-121">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-121">Compose</span></span><br><span data-ttu-id="d72d0-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-122">Read</span></span> | [<span data-ttu-id="d72d0-123">Auth</span><span class="sxs-lookup"><span data-stu-id="d72d0-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-124">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d72d0-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d72d0-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="d72d0-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="d72d0-126">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-126">Compose</span></span><br><span data-ttu-id="d72d0-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-127">Read</span></span> | <span data-ttu-id="d72d0-128">文字列</span><span class="sxs-lookup"><span data-stu-id="d72d0-128">String</span></span> | [<span data-ttu-id="d72d0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-130">ダン</span><span class="sxs-lookup"><span data-stu-id="d72d0-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="d72d0-131">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-131">Compose</span></span><br><span data-ttu-id="d72d0-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-132">Read</span></span> | [<span data-ttu-id="d72d0-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d72d0-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="d72d0-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="d72d0-136">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-136">Compose</span></span><br><span data-ttu-id="d72d0-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-137">Read</span></span> | <span data-ttu-id="d72d0-138">文字列</span><span class="sxs-lookup"><span data-stu-id="d72d0-138">String</span></span> | [<span data-ttu-id="d72d0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-140">主催</span><span class="sxs-lookup"><span data-stu-id="d72d0-140">host</span></span>](#host-hosttype) | <span data-ttu-id="d72d0-141">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-141">Compose</span></span><br><span data-ttu-id="d72d0-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-142">Read</span></span> | [<span data-ttu-id="d72d0-143">HostType</span><span class="sxs-lookup"><span data-stu-id="d72d0-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="d72d0-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="d72d0-146">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-146">Compose</span></span><br><span data-ttu-id="d72d0-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-147">Read</span></span> | [<span data-ttu-id="d72d0-148">メールボックス</span><span class="sxs-lookup"><span data-stu-id="d72d0-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="d72d0-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="d72d0-151">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-151">Compose</span></span><br><span data-ttu-id="d72d0-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-152">Read</span></span> | [<span data-ttu-id="d72d0-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d72d0-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-154">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d72d0-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d72d0-155">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="d72d0-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="d72d0-156">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-156">Compose</span></span><br><span data-ttu-id="d72d0-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-157">Read</span></span> | [<span data-ttu-id="d72d0-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d72d0-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-160">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="d72d0-161">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-161">Compose</span></span><br><span data-ttu-id="d72d0-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-162">Read</span></span> | [<span data-ttu-id="d72d0-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d72d0-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="d72d0-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="d72d0-166">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-166">Compose</span></span><br><span data-ttu-id="d72d0-167">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-167">Read</span></span> | [<span data-ttu-id="d72d0-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d72d0-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d72d0-170">UI</span><span class="sxs-lookup"><span data-stu-id="d72d0-170">ui</span></span>](#ui-ui) | <span data-ttu-id="d72d0-171">作成</span><span class="sxs-lookup"><span data-stu-id="d72d0-171">Compose</span></span><br><span data-ttu-id="d72d0-172">読み取り</span><span class="sxs-lookup"><span data-stu-id="d72d0-172">Read</span></span> | [<span data-ttu-id="d72d0-173">UI</span><span class="sxs-lookup"><span data-stu-id="d72d0-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d72d0-174">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="d72d0-175">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="d72d0-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="d72d0-176">auth: [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="d72d0-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="d72d0-177">[シングルサインオン (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md)をサポートするために、Office アプリケーションがアドインの web アプリケーションへのアクセストークンを取得できるようにする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="d72d0-178">これにより、間接的に、サインインしたユーザーの Microsoft Graph データにアドインがアクセスできるようにもなります。ユーザーがもう一度サインインする必要はありません。</span><span class="sxs-lookup"><span data-stu-id="d72d0-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-179">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-179">Type</span></span>

*   [<span data-ttu-id="d72d0-180">Auth</span><span class="sxs-lookup"><span data-stu-id="d72d0-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="d72d0-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-181">Requirements</span></span>

|<span data-ttu-id="d72d0-182">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-182">Requirement</span></span>| <span data-ttu-id="d72d0-183">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-184">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-185">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d72d0-185">Preview</span></span>|
|[<span data-ttu-id="d72d0-186">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-187">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-188">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="d72d0-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="d72d0-189">contentLanguage: String</span></span>

<span data-ttu-id="d72d0-190">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="d72d0-191">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-192">型</span><span class="sxs-lookup"><span data-stu-id="d72d0-192">Type</span></span>

*   <span data-ttu-id="d72d0-193">String</span><span class="sxs-lookup"><span data-stu-id="d72d0-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d72d0-194">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-194">Requirements</span></span>

|<span data-ttu-id="d72d0-195">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-195">Requirement</span></span>| <span data-ttu-id="d72d0-196">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-197">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-198">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-198">1.1</span></span>|
|[<span data-ttu-id="d72d0-199">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-200">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-201">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="d72d0-202">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="d72d0-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="d72d0-203">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-204">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-204">Type</span></span>

*   [<span data-ttu-id="d72d0-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d72d0-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="d72d0-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-206">Requirements</span></span>

|<span data-ttu-id="d72d0-207">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-207">Requirement</span></span>| <span data-ttu-id="d72d0-208">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-209">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-210">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-210">1.1</span></span>|
|[<span data-ttu-id="d72d0-211">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-212">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-213">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="d72d0-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="d72d0-214">displayLanguage: String</span></span>

<span data-ttu-id="d72d0-215">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="d72d0-216">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-217">型</span><span class="sxs-lookup"><span data-stu-id="d72d0-217">Type</span></span>

*   <span data-ttu-id="d72d0-218">String</span><span class="sxs-lookup"><span data-stu-id="d72d0-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d72d0-219">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-219">Requirements</span></span>

|<span data-ttu-id="d72d0-220">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-220">Requirement</span></span>| <span data-ttu-id="d72d0-221">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-223">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-223">1.1</span></span>|
|[<span data-ttu-id="d72d0-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-226">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="d72d0-227">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="d72d0-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="d72d0-228">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-229">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-229">Type</span></span>

*   [<span data-ttu-id="d72d0-230">HostType</span><span class="sxs-lookup"><span data-stu-id="d72d0-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="d72d0-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-231">Requirements</span></span>

|<span data-ttu-id="d72d0-232">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-232">Requirement</span></span>| <span data-ttu-id="d72d0-233">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-235">1.1</span></span>|
|[<span data-ttu-id="d72d0-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-238">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="d72d0-239">officeTheme: [Officetheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="d72d0-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="d72d0-240">Office テーマの色のプロパティにアクセスできるようにします。</span><span class="sxs-lookup"><span data-stu-id="d72d0-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="d72d0-241">このメンバーは、Windows の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="d72d0-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="d72d0-242">Office テーマの色を使用すると、アドインの配色を、[ **ファイル > Office アカウント > Office テーマ UI**を使用してユーザーが選択した現在の office テーマを使用して調整できます。これは、すべての office クライアントアプリケーションで適用されます。</span><span class="sxs-lookup"><span data-stu-id="d72d0-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="d72d0-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="d72d0-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-244">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-244">Type</span></span>

*   [<span data-ttu-id="d72d0-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d72d0-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="d72d0-246">プロパティ:</span><span class="sxs-lookup"><span data-stu-id="d72d0-246">Properties:</span></span>

|<span data-ttu-id="d72d0-247">名前</span><span class="sxs-lookup"><span data-stu-id="d72d0-247">Name</span></span>| <span data-ttu-id="d72d0-248">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-248">Type</span></span>| <span data-ttu-id="d72d0-249">説明</span><span class="sxs-lookup"><span data-stu-id="d72d0-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="d72d0-250">文字列</span><span class="sxs-lookup"><span data-stu-id="d72d0-250">String</span></span>|<span data-ttu-id="d72d0-251">Office テーマの本文の背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="d72d0-252">String</span><span class="sxs-lookup"><span data-stu-id="d72d0-252">String</span></span>|<span data-ttu-id="d72d0-253">Office テーマの本文の前景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="d72d0-254">String</span><span class="sxs-lookup"><span data-stu-id="d72d0-254">String</span></span>|<span data-ttu-id="d72d0-255">Office テーマのコントロールの背景色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="d72d0-256">String</span><span class="sxs-lookup"><span data-stu-id="d72d0-256">String</span></span>|<span data-ttu-id="d72d0-257">Office テーマの本文のコントロール色を 16 進数の組み合わせとして取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d72d0-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-258">Requirements</span></span>

|<span data-ttu-id="d72d0-259">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-259">Requirement</span></span>| <span data-ttu-id="d72d0-260">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-262">プレビュー</span><span class="sxs-lookup"><span data-stu-id="d72d0-262">Preview</span></span>|
|[<span data-ttu-id="d72d0-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-265">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="d72d0-266">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="d72d0-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="d72d0-267">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-268">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-268">Type</span></span>

*   [<span data-ttu-id="d72d0-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d72d0-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="d72d0-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-270">Requirements</span></span>

|<span data-ttu-id="d72d0-271">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-271">Requirement</span></span>| <span data-ttu-id="d72d0-272">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-273">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-274">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-274">1.1</span></span>|
|[<span data-ttu-id="d72d0-275">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-276">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-277">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="d72d0-278">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="d72d0-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="d72d0-279">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-280">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-280">Type</span></span>

*   [<span data-ttu-id="d72d0-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d72d0-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="d72d0-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-282">Requirements</span></span>

|<span data-ttu-id="d72d0-283">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-283">Requirement</span></span>| <span data-ttu-id="d72d0-284">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-285">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-286">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-286">1.1</span></span>|
|[<span data-ttu-id="d72d0-287">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-288">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d72d0-289">例</span><span class="sxs-lookup"><span data-stu-id="d72d0-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="d72d0-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="d72d0-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="d72d0-291">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d72d0-292">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="d72d0-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-293">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-293">Type</span></span>

*   [<span data-ttu-id="d72d0-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d72d0-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="d72d0-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-295">Requirements</span></span>

|<span data-ttu-id="d72d0-296">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-296">Requirement</span></span>| <span data-ttu-id="d72d0-297">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-298">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-299">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-299">1.1</span></span>|
|[<span data-ttu-id="d72d0-300">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="d72d0-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="d72d0-301">制限あり</span><span class="sxs-lookup"><span data-stu-id="d72d0-301">Restricted</span></span>|
|[<span data-ttu-id="d72d0-302">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-303">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="d72d0-304">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="d72d0-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="d72d0-305">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="d72d0-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d72d0-306">種類</span><span class="sxs-lookup"><span data-stu-id="d72d0-306">Type</span></span>

*   [<span data-ttu-id="d72d0-307">UI</span><span class="sxs-lookup"><span data-stu-id="d72d0-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="d72d0-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="d72d0-308">Requirements</span></span>

|<span data-ttu-id="d72d0-309">要件</span><span class="sxs-lookup"><span data-stu-id="d72d0-309">Requirement</span></span>| <span data-ttu-id="d72d0-310">値</span><span class="sxs-lookup"><span data-stu-id="d72d0-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d72d0-311">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="d72d0-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d72d0-312">1.1</span><span class="sxs-lookup"><span data-stu-id="d72d0-312">1.1</span></span>|
|[<span data-ttu-id="d72d0-313">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="d72d0-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d72d0-314">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="d72d0-314">Compose or Read</span></span>|
