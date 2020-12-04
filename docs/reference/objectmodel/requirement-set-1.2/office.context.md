---
title: Office コンテキスト要件セット1.2
description: メールボックス API 要件セット1.2 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 1b697cbe29be7d0af6fec65e47d080ebd1af17ae
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570780"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="1065c-103">コンテキスト (メールボックス要件セット 1.2)</span><span class="sxs-lookup"><span data-stu-id="1065c-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="1065c-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1065c-104">[Office](office.md).context</span></span>

<span data-ttu-id="1065c-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1065c-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1065c-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1065c-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1065c-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="1065c-107">Requirements</span></span>

|<span data-ttu-id="1065c-108">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-108">Requirement</span></span>| <span data-ttu-id="1065c-109">値</span><span class="sxs-lookup"><span data-stu-id="1065c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-111">1.1</span></span>|
|[<span data-ttu-id="1065c-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1065c-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1065c-114">Properties</span></span>

| <span data-ttu-id="1065c-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1065c-115">Property</span></span> | <span data-ttu-id="1065c-116">モード</span><span class="sxs-lookup"><span data-stu-id="1065c-116">Modes</span></span> | <span data-ttu-id="1065c-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1065c-117">Return type</span></span> | <span data-ttu-id="1065c-118">最小値</span><span class="sxs-lookup"><span data-stu-id="1065c-118">Minimum</span></span><br><span data-ttu-id="1065c-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="1065c-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1065c-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1065c-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1065c-121">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-121">Compose</span></span><br><span data-ttu-id="1065c-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-122">Read</span></span> | <span data-ttu-id="1065c-123">文字列</span><span class="sxs-lookup"><span data-stu-id="1065c-123">String</span></span> | [<span data-ttu-id="1065c-124">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-125">ダン</span><span class="sxs-lookup"><span data-stu-id="1065c-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1065c-126">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-126">Compose</span></span><br><span data-ttu-id="1065c-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-127">Read</span></span> | [<span data-ttu-id="1065c-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1065c-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="1065c-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1065c-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1065c-131">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-131">Compose</span></span><br><span data-ttu-id="1065c-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-132">Read</span></span> | <span data-ttu-id="1065c-133">文字列</span><span class="sxs-lookup"><span data-stu-id="1065c-133">String</span></span> | [<span data-ttu-id="1065c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="1065c-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1065c-136">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-136">Compose</span></span><br><span data-ttu-id="1065c-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-137">Read</span></span> | [<span data-ttu-id="1065c-138">メールボックス</span><span class="sxs-lookup"><span data-stu-id="1065c-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="1065c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-140">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1065c-141">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-141">Compose</span></span><br><span data-ttu-id="1065c-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-142">Read</span></span> | [<span data-ttu-id="1065c-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1065c-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="1065c-144">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1065c-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1065c-146">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-146">Compose</span></span><br><span data-ttu-id="1065c-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-147">Read</span></span> | [<span data-ttu-id="1065c-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1065c-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="1065c-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1065c-150">UI</span><span class="sxs-lookup"><span data-stu-id="1065c-150">ui</span></span>](#ui-ui) | <span data-ttu-id="1065c-151">作成</span><span class="sxs-lookup"><span data-stu-id="1065c-151">Compose</span></span><br><span data-ttu-id="1065c-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="1065c-152">Read</span></span> | [<span data-ttu-id="1065c-153">UI</span><span class="sxs-lookup"><span data-stu-id="1065c-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="1065c-154">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1065c-155">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="1065c-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1065c-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1065c-156">contentLanguage: String</span></span>

<span data-ttu-id="1065c-157">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="1065c-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1065c-158">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="1065c-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-159">型</span><span class="sxs-lookup"><span data-stu-id="1065c-159">Type</span></span>

*   <span data-ttu-id="1065c-160">String</span><span class="sxs-lookup"><span data-stu-id="1065c-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1065c-161">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-161">Requirements</span></span>

|<span data-ttu-id="1065c-162">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-162">Requirement</span></span>| <span data-ttu-id="1065c-163">値</span><span class="sxs-lookup"><span data-stu-id="1065c-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-165">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-165">1.1</span></span>|
|[<span data-ttu-id="1065c-166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-167">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1065c-168">例</span><span class="sxs-lookup"><span data-stu-id="1065c-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1065c-169">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1065c-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1065c-170">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="1065c-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-171">種類</span><span class="sxs-lookup"><span data-stu-id="1065c-171">Type</span></span>

*   [<span data-ttu-id="1065c-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1065c-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1065c-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="1065c-173">Requirements</span></span>

|<span data-ttu-id="1065c-174">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-174">Requirement</span></span>| <span data-ttu-id="1065c-175">値</span><span class="sxs-lookup"><span data-stu-id="1065c-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-177">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-177">1.1</span></span>|
|[<span data-ttu-id="1065c-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1065c-180">例</span><span class="sxs-lookup"><span data-stu-id="1065c-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1065c-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1065c-181">displayLanguage: String</span></span>

<span data-ttu-id="1065c-182">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="1065c-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="1065c-183">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="1065c-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-184">型</span><span class="sxs-lookup"><span data-stu-id="1065c-184">Type</span></span>

*   <span data-ttu-id="1065c-185">String</span><span class="sxs-lookup"><span data-stu-id="1065c-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1065c-186">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-186">Requirements</span></span>

|<span data-ttu-id="1065c-187">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-187">Requirement</span></span>| <span data-ttu-id="1065c-188">値</span><span class="sxs-lookup"><span data-stu-id="1065c-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-190">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-190">1.1</span></span>|
|[<span data-ttu-id="1065c-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1065c-193">例</span><span class="sxs-lookup"><span data-stu-id="1065c-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1065c-194">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1065c-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1065c-195">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="1065c-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-196">種類</span><span class="sxs-lookup"><span data-stu-id="1065c-196">Type</span></span>

*   [<span data-ttu-id="1065c-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1065c-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1065c-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="1065c-198">Requirements</span></span>

|<span data-ttu-id="1065c-199">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-199">Requirement</span></span>| <span data-ttu-id="1065c-200">値</span><span class="sxs-lookup"><span data-stu-id="1065c-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-202">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-202">1.1</span></span>|
|[<span data-ttu-id="1065c-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1065c-205">例</span><span class="sxs-lookup"><span data-stu-id="1065c-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1065c-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1065c-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1065c-207">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1065c-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1065c-208">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="1065c-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-209">種類</span><span class="sxs-lookup"><span data-stu-id="1065c-209">Type</span></span>

*   [<span data-ttu-id="1065c-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1065c-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1065c-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="1065c-211">Requirements</span></span>

|<span data-ttu-id="1065c-212">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-212">Requirement</span></span>| <span data-ttu-id="1065c-213">値</span><span class="sxs-lookup"><span data-stu-id="1065c-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-215">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-215">1.1</span></span>|
|[<span data-ttu-id="1065c-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1065c-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1065c-217">制限あり</span><span class="sxs-lookup"><span data-stu-id="1065c-217">Restricted</span></span>|
|[<span data-ttu-id="1065c-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-219">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1065c-220">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1065c-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1065c-221">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="1065c-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1065c-222">種類</span><span class="sxs-lookup"><span data-stu-id="1065c-222">Type</span></span>

*   [<span data-ttu-id="1065c-223">UI</span><span class="sxs-lookup"><span data-stu-id="1065c-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1065c-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="1065c-224">Requirements</span></span>

|<span data-ttu-id="1065c-225">要件</span><span class="sxs-lookup"><span data-stu-id="1065c-225">Requirement</span></span>| <span data-ttu-id="1065c-226">値</span><span class="sxs-lookup"><span data-stu-id="1065c-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="1065c-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1065c-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1065c-228">1.1</span><span class="sxs-lookup"><span data-stu-id="1065c-228">1.1</span></span>|
|[<span data-ttu-id="1065c-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1065c-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1065c-230">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1065c-230">Compose or Read</span></span>|
