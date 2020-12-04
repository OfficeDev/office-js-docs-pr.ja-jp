---
title: Office コンテキスト要件セット1.3
description: メールボックス API 要件セット1.3 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: b497cdf3f878df7efd816f236bd565c8fad7d922
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570745"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="6b9d3-103">コンテキスト (メールボックス要件セット 1.3)</span><span class="sxs-lookup"><span data-stu-id="6b9d3-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6b9d3-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6b9d3-104">[Office](office.md).context</span></span>

<span data-ttu-id="6b9d3-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6b9d3-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b9d3-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="6b9d3-107">Requirements</span></span>

|<span data-ttu-id="6b9d3-108">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-108">Requirement</span></span>| <span data-ttu-id="6b9d3-109">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-111">1.1</span></span>|
|[<span data-ttu-id="6b9d3-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6b9d3-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6b9d3-114">Properties</span></span>

| <span data-ttu-id="6b9d3-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6b9d3-115">Property</span></span> | <span data-ttu-id="6b9d3-116">モード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-116">Modes</span></span> | <span data-ttu-id="6b9d3-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="6b9d3-117">Return type</span></span> | <span data-ttu-id="6b9d3-118">最小値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-118">Minimum</span></span><br><span data-ttu-id="6b9d3-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="6b9d3-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6b9d3-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6b9d3-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6b9d3-121">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-121">Compose</span></span><br><span data-ttu-id="6b9d3-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-122">Read</span></span> | <span data-ttu-id="6b9d3-123">文字列</span><span class="sxs-lookup"><span data-stu-id="6b9d3-123">String</span></span> | [<span data-ttu-id="6b9d3-124">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-125">ダン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6b9d3-126">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-126">Compose</span></span><br><span data-ttu-id="6b9d3-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-127">Read</span></span> | [<span data-ttu-id="6b9d3-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6b9d3-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="6b9d3-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6b9d3-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6b9d3-131">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-131">Compose</span></span><br><span data-ttu-id="6b9d3-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-132">Read</span></span> | <span data-ttu-id="6b9d3-133">文字列</span><span class="sxs-lookup"><span data-stu-id="6b9d3-133">String</span></span> | [<span data-ttu-id="6b9d3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="6b9d3-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6b9d3-136">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-136">Compose</span></span><br><span data-ttu-id="6b9d3-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-137">Read</span></span> | [<span data-ttu-id="6b9d3-138">メールボックス</span><span class="sxs-lookup"><span data-stu-id="6b9d3-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="6b9d3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-140">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6b9d3-141">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-141">Compose</span></span><br><span data-ttu-id="6b9d3-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-142">Read</span></span> | [<span data-ttu-id="6b9d3-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6b9d3-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="6b9d3-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6b9d3-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6b9d3-146">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-146">Compose</span></span><br><span data-ttu-id="6b9d3-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-147">Read</span></span> | [<span data-ttu-id="6b9d3-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6b9d3-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="6b9d3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6b9d3-150">UI</span><span class="sxs-lookup"><span data-stu-id="6b9d3-150">ui</span></span>](#ui-ui) | <span data-ttu-id="6b9d3-151">作成</span><span class="sxs-lookup"><span data-stu-id="6b9d3-151">Compose</span></span><br><span data-ttu-id="6b9d3-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="6b9d3-152">Read</span></span> | [<span data-ttu-id="6b9d3-153">UI</span><span class="sxs-lookup"><span data-stu-id="6b9d3-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="6b9d3-154">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6b9d3-155">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="6b9d3-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="6b9d3-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6b9d3-156">contentLanguage: String</span></span>

<span data-ttu-id="6b9d3-157">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6b9d3-158">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-159">型</span><span class="sxs-lookup"><span data-stu-id="6b9d3-159">Type</span></span>

*   <span data-ttu-id="6b9d3-160">String</span><span class="sxs-lookup"><span data-stu-id="6b9d3-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b9d3-161">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-161">Requirements</span></span>

|<span data-ttu-id="6b9d3-162">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-162">Requirement</span></span>| <span data-ttu-id="6b9d3-163">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-165">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-165">1.1</span></span>|
|[<span data-ttu-id="6b9d3-166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-167">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b9d3-168">例</span><span class="sxs-lookup"><span data-stu-id="6b9d3-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6b9d3-169">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6b9d3-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6b9d3-170">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-171">種類</span><span class="sxs-lookup"><span data-stu-id="6b9d3-171">Type</span></span>

*   [<span data-ttu-id="6b9d3-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6b9d3-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6b9d3-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="6b9d3-173">Requirements</span></span>

|<span data-ttu-id="6b9d3-174">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-174">Requirement</span></span>| <span data-ttu-id="6b9d3-175">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-177">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-177">1.1</span></span>|
|[<span data-ttu-id="6b9d3-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b9d3-180">例</span><span class="sxs-lookup"><span data-stu-id="6b9d3-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6b9d3-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6b9d3-181">displayLanguage: String</span></span>

<span data-ttu-id="6b9d3-182">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="6b9d3-183">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-184">型</span><span class="sxs-lookup"><span data-stu-id="6b9d3-184">Type</span></span>

*   <span data-ttu-id="6b9d3-185">String</span><span class="sxs-lookup"><span data-stu-id="6b9d3-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b9d3-186">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-186">Requirements</span></span>

|<span data-ttu-id="6b9d3-187">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-187">Requirement</span></span>| <span data-ttu-id="6b9d3-188">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-190">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-190">1.1</span></span>|
|[<span data-ttu-id="6b9d3-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b9d3-193">例</span><span class="sxs-lookup"><span data-stu-id="6b9d3-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6b9d3-194">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6b9d3-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6b9d3-195">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-196">種類</span><span class="sxs-lookup"><span data-stu-id="6b9d3-196">Type</span></span>

*   [<span data-ttu-id="6b9d3-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6b9d3-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6b9d3-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="6b9d3-198">Requirements</span></span>

|<span data-ttu-id="6b9d3-199">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-199">Requirement</span></span>| <span data-ttu-id="6b9d3-200">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-202">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-202">1.1</span></span>|
|[<span data-ttu-id="6b9d3-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b9d3-205">例</span><span class="sxs-lookup"><span data-stu-id="6b9d3-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6b9d3-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6b9d3-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6b9d3-207">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6b9d3-208">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-209">種類</span><span class="sxs-lookup"><span data-stu-id="6b9d3-209">Type</span></span>

*   [<span data-ttu-id="6b9d3-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6b9d3-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6b9d3-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="6b9d3-211">Requirements</span></span>

|<span data-ttu-id="6b9d3-212">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-212">Requirement</span></span>| <span data-ttu-id="6b9d3-213">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-215">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-215">1.1</span></span>|
|[<span data-ttu-id="6b9d3-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6b9d3-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6b9d3-217">制限あり</span><span class="sxs-lookup"><span data-stu-id="6b9d3-217">Restricted</span></span>|
|[<span data-ttu-id="6b9d3-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-219">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6b9d3-220">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6b9d3-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6b9d3-221">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6b9d3-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6b9d3-222">種類</span><span class="sxs-lookup"><span data-stu-id="6b9d3-222">Type</span></span>

*   [<span data-ttu-id="6b9d3-223">UI</span><span class="sxs-lookup"><span data-stu-id="6b9d3-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6b9d3-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="6b9d3-224">Requirements</span></span>

|<span data-ttu-id="6b9d3-225">要件</span><span class="sxs-lookup"><span data-stu-id="6b9d3-225">Requirement</span></span>| <span data-ttu-id="6b9d3-226">値</span><span class="sxs-lookup"><span data-stu-id="6b9d3-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b9d3-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6b9d3-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6b9d3-228">1.1</span><span class="sxs-lookup"><span data-stu-id="6b9d3-228">1.1</span></span>|
|[<span data-ttu-id="6b9d3-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6b9d3-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6b9d3-230">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6b9d3-230">Compose or Read</span></span>|
