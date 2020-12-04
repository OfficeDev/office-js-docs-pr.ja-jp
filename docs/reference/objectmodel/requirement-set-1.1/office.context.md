---
title: Office コンテキスト要件セット1.1
description: メールボックス API 要件セット1.1 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 4f0fa4094477125f4d07fd6ddb4ac2c3c08a5d70
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570752"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="83d5e-103">コンテキスト (メールボックス要件セット 1.1)</span><span class="sxs-lookup"><span data-stu-id="83d5e-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="83d5e-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="83d5e-104">[Office](office.md).context</span></span>

<span data-ttu-id="83d5e-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="83d5e-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="83d5e-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="83d5e-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="83d5e-107">Requirements</span></span>

|<span data-ttu-id="83d5e-108">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-108">Requirement</span></span>| <span data-ttu-id="83d5e-109">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-111">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-111">1.1</span></span>|
|[<span data-ttu-id="83d5e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="83d5e-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="83d5e-114">Properties</span></span>

| <span data-ttu-id="83d5e-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="83d5e-115">Property</span></span> | <span data-ttu-id="83d5e-116">モード</span><span class="sxs-lookup"><span data-stu-id="83d5e-116">Modes</span></span> | <span data-ttu-id="83d5e-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="83d5e-117">Return type</span></span> | <span data-ttu-id="83d5e-118">最小値</span><span class="sxs-lookup"><span data-stu-id="83d5e-118">Minimum</span></span><br><span data-ttu-id="83d5e-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="83d5e-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="83d5e-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="83d5e-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="83d5e-121">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-121">Compose</span></span><br><span data-ttu-id="83d5e-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-122">Read</span></span> | <span data-ttu-id="83d5e-123">文字列</span><span class="sxs-lookup"><span data-stu-id="83d5e-123">String</span></span> | [<span data-ttu-id="83d5e-124">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-125">ダン</span><span class="sxs-lookup"><span data-stu-id="83d5e-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="83d5e-126">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-126">Compose</span></span><br><span data-ttu-id="83d5e-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-127">Read</span></span> | [<span data-ttu-id="83d5e-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="83d5e-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="83d5e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="83d5e-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="83d5e-131">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-131">Compose</span></span><br><span data-ttu-id="83d5e-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-132">Read</span></span> | <span data-ttu-id="83d5e-133">文字列</span><span class="sxs-lookup"><span data-stu-id="83d5e-133">String</span></span> | [<span data-ttu-id="83d5e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="83d5e-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="83d5e-136">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-136">Compose</span></span><br><span data-ttu-id="83d5e-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-137">Read</span></span> | [<span data-ttu-id="83d5e-138">メールボックス</span><span class="sxs-lookup"><span data-stu-id="83d5e-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="83d5e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-140">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="83d5e-141">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-141">Compose</span></span><br><span data-ttu-id="83d5e-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-142">Read</span></span> | [<span data-ttu-id="83d5e-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="83d5e-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="83d5e-144">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="83d5e-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="83d5e-146">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-146">Compose</span></span><br><span data-ttu-id="83d5e-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-147">Read</span></span> | [<span data-ttu-id="83d5e-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="83d5e-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="83d5e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="83d5e-150">UI</span><span class="sxs-lookup"><span data-stu-id="83d5e-150">ui</span></span>](#ui-ui) | <span data-ttu-id="83d5e-151">作成</span><span class="sxs-lookup"><span data-stu-id="83d5e-151">Compose</span></span><br><span data-ttu-id="83d5e-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="83d5e-152">Read</span></span> | [<span data-ttu-id="83d5e-153">UI</span><span class="sxs-lookup"><span data-stu-id="83d5e-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="83d5e-154">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="83d5e-155">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="83d5e-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="83d5e-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="83d5e-156">contentLanguage: String</span></span>

<span data-ttu-id="83d5e-157">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="83d5e-158">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-159">型</span><span class="sxs-lookup"><span data-stu-id="83d5e-159">Type</span></span>

*   <span data-ttu-id="83d5e-160">String</span><span class="sxs-lookup"><span data-stu-id="83d5e-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83d5e-161">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-161">Requirements</span></span>

|<span data-ttu-id="83d5e-162">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-162">Requirement</span></span>| <span data-ttu-id="83d5e-163">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-165">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-165">1.1</span></span>|
|[<span data-ttu-id="83d5e-166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-167">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83d5e-168">例</span><span class="sxs-lookup"><span data-stu-id="83d5e-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="83d5e-169">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="83d5e-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="83d5e-170">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-171">種類</span><span class="sxs-lookup"><span data-stu-id="83d5e-171">Type</span></span>

*   [<span data-ttu-id="83d5e-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="83d5e-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="83d5e-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="83d5e-173">Requirements</span></span>

|<span data-ttu-id="83d5e-174">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-174">Requirement</span></span>| <span data-ttu-id="83d5e-175">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-177">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-177">1.1</span></span>|
|[<span data-ttu-id="83d5e-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83d5e-180">例</span><span class="sxs-lookup"><span data-stu-id="83d5e-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="83d5e-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="83d5e-181">displayLanguage: String</span></span>

<span data-ttu-id="83d5e-182">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="83d5e-183">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-184">型</span><span class="sxs-lookup"><span data-stu-id="83d5e-184">Type</span></span>

*   <span data-ttu-id="83d5e-185">String</span><span class="sxs-lookup"><span data-stu-id="83d5e-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="83d5e-186">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-186">Requirements</span></span>

|<span data-ttu-id="83d5e-187">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-187">Requirement</span></span>| <span data-ttu-id="83d5e-188">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-190">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-190">1.1</span></span>|
|[<span data-ttu-id="83d5e-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83d5e-193">例</span><span class="sxs-lookup"><span data-stu-id="83d5e-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="83d5e-194">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="83d5e-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="83d5e-195">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-196">種類</span><span class="sxs-lookup"><span data-stu-id="83d5e-196">Type</span></span>

*   [<span data-ttu-id="83d5e-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="83d5e-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="83d5e-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="83d5e-198">Requirements</span></span>

|<span data-ttu-id="83d5e-199">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-199">Requirement</span></span>| <span data-ttu-id="83d5e-200">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-202">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-202">1.1</span></span>|
|[<span data-ttu-id="83d5e-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="83d5e-205">例</span><span class="sxs-lookup"><span data-stu-id="83d5e-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="83d5e-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="83d5e-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="83d5e-207">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="83d5e-208">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="83d5e-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-209">種類</span><span class="sxs-lookup"><span data-stu-id="83d5e-209">Type</span></span>

*   [<span data-ttu-id="83d5e-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="83d5e-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="83d5e-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="83d5e-211">Requirements</span></span>

|<span data-ttu-id="83d5e-212">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-212">Requirement</span></span>| <span data-ttu-id="83d5e-213">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-215">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-215">1.1</span></span>|
|[<span data-ttu-id="83d5e-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="83d5e-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="83d5e-217">制限あり</span><span class="sxs-lookup"><span data-stu-id="83d5e-217">Restricted</span></span>|
|[<span data-ttu-id="83d5e-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-219">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="83d5e-220">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="83d5e-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="83d5e-221">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="83d5e-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="83d5e-222">種類</span><span class="sxs-lookup"><span data-stu-id="83d5e-222">Type</span></span>

*   [<span data-ttu-id="83d5e-223">UI</span><span class="sxs-lookup"><span data-stu-id="83d5e-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="83d5e-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="83d5e-224">Requirements</span></span>

|<span data-ttu-id="83d5e-225">要件</span><span class="sxs-lookup"><span data-stu-id="83d5e-225">Requirement</span></span>| <span data-ttu-id="83d5e-226">値</span><span class="sxs-lookup"><span data-stu-id="83d5e-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="83d5e-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="83d5e-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="83d5e-228">1.1</span><span class="sxs-lookup"><span data-stu-id="83d5e-228">1.1</span></span>|
|[<span data-ttu-id="83d5e-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="83d5e-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="83d5e-230">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="83d5e-230">Compose or Read</span></span>|
