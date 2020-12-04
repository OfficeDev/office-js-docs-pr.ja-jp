---
title: Office コンテキスト要件セット1.4
description: メールボックス API 要件セット1.4 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 0ec84c9d0695871fa3be265c37ce1e682cdfb6af
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570773"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="eef7d-103">コンテキスト (メールボックス要件セット 1.4)</span><span class="sxs-lookup"><span data-stu-id="eef7d-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="eef7d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="eef7d-104">[Office](office.md).context</span></span>

<span data-ttu-id="eef7d-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="eef7d-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="eef7d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eef7d-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="eef7d-107">Requirements</span></span>

|<span data-ttu-id="eef7d-108">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-108">Requirement</span></span>| <span data-ttu-id="eef7d-109">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-111">1.1</span></span>|
|[<span data-ttu-id="eef7d-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="eef7d-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="eef7d-114">Properties</span></span>

| <span data-ttu-id="eef7d-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="eef7d-115">Property</span></span> | <span data-ttu-id="eef7d-116">モード</span><span class="sxs-lookup"><span data-stu-id="eef7d-116">Modes</span></span> | <span data-ttu-id="eef7d-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="eef7d-117">Return type</span></span> | <span data-ttu-id="eef7d-118">最小値</span><span class="sxs-lookup"><span data-stu-id="eef7d-118">Minimum</span></span><br><span data-ttu-id="eef7d-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="eef7d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="eef7d-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="eef7d-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="eef7d-121">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-121">Compose</span></span><br><span data-ttu-id="eef7d-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-122">Read</span></span> | <span data-ttu-id="eef7d-123">文字列</span><span class="sxs-lookup"><span data-stu-id="eef7d-123">String</span></span> | [<span data-ttu-id="eef7d-124">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-125">ダン</span><span class="sxs-lookup"><span data-stu-id="eef7d-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="eef7d-126">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-126">Compose</span></span><br><span data-ttu-id="eef7d-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-127">Read</span></span> | [<span data-ttu-id="eef7d-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="eef7d-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="eef7d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="eef7d-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="eef7d-131">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-131">Compose</span></span><br><span data-ttu-id="eef7d-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-132">Read</span></span> | <span data-ttu-id="eef7d-133">文字列</span><span class="sxs-lookup"><span data-stu-id="eef7d-133">String</span></span> | [<span data-ttu-id="eef7d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="eef7d-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="eef7d-136">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-136">Compose</span></span><br><span data-ttu-id="eef7d-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-137">Read</span></span> | [<span data-ttu-id="eef7d-138">メールボックス</span><span class="sxs-lookup"><span data-stu-id="eef7d-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="eef7d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-140">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="eef7d-141">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-141">Compose</span></span><br><span data-ttu-id="eef7d-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-142">Read</span></span> | [<span data-ttu-id="eef7d-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="eef7d-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="eef7d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="eef7d-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="eef7d-146">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-146">Compose</span></span><br><span data-ttu-id="eef7d-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-147">Read</span></span> | [<span data-ttu-id="eef7d-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eef7d-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="eef7d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eef7d-150">UI</span><span class="sxs-lookup"><span data-stu-id="eef7d-150">ui</span></span>](#ui-ui) | <span data-ttu-id="eef7d-151">作成</span><span class="sxs-lookup"><span data-stu-id="eef7d-151">Compose</span></span><br><span data-ttu-id="eef7d-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="eef7d-152">Read</span></span> | [<span data-ttu-id="eef7d-153">UI</span><span class="sxs-lookup"><span data-stu-id="eef7d-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="eef7d-154">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="eef7d-155">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="eef7d-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="eef7d-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="eef7d-156">contentLanguage: String</span></span>

<span data-ttu-id="eef7d-157">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="eef7d-158">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-159">型</span><span class="sxs-lookup"><span data-stu-id="eef7d-159">Type</span></span>

*   <span data-ttu-id="eef7d-160">String</span><span class="sxs-lookup"><span data-stu-id="eef7d-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eef7d-161">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-161">Requirements</span></span>

|<span data-ttu-id="eef7d-162">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-162">Requirement</span></span>| <span data-ttu-id="eef7d-163">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-165">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-165">1.1</span></span>|
|[<span data-ttu-id="eef7d-166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-167">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eef7d-168">例</span><span class="sxs-lookup"><span data-stu-id="eef7d-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="eef7d-169">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="eef7d-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="eef7d-170">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-171">種類</span><span class="sxs-lookup"><span data-stu-id="eef7d-171">Type</span></span>

*   [<span data-ttu-id="eef7d-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="eef7d-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="eef7d-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="eef7d-173">Requirements</span></span>

|<span data-ttu-id="eef7d-174">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-174">Requirement</span></span>| <span data-ttu-id="eef7d-175">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-177">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-177">1.1</span></span>|
|[<span data-ttu-id="eef7d-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eef7d-180">例</span><span class="sxs-lookup"><span data-stu-id="eef7d-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="eef7d-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="eef7d-181">displayLanguage: String</span></span>

<span data-ttu-id="eef7d-182">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="eef7d-183">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-184">型</span><span class="sxs-lookup"><span data-stu-id="eef7d-184">Type</span></span>

*   <span data-ttu-id="eef7d-185">String</span><span class="sxs-lookup"><span data-stu-id="eef7d-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eef7d-186">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-186">Requirements</span></span>

|<span data-ttu-id="eef7d-187">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-187">Requirement</span></span>| <span data-ttu-id="eef7d-188">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-190">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-190">1.1</span></span>|
|[<span data-ttu-id="eef7d-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eef7d-193">例</span><span class="sxs-lookup"><span data-stu-id="eef7d-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="eef7d-194">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="eef7d-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="eef7d-195">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-196">種類</span><span class="sxs-lookup"><span data-stu-id="eef7d-196">Type</span></span>

*   [<span data-ttu-id="eef7d-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="eef7d-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="eef7d-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="eef7d-198">Requirements</span></span>

|<span data-ttu-id="eef7d-199">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-199">Requirement</span></span>| <span data-ttu-id="eef7d-200">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-202">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-202">1.1</span></span>|
|[<span data-ttu-id="eef7d-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eef7d-205">例</span><span class="sxs-lookup"><span data-stu-id="eef7d-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="eef7d-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="eef7d-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="eef7d-207">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="eef7d-208">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="eef7d-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-209">種類</span><span class="sxs-lookup"><span data-stu-id="eef7d-209">Type</span></span>

*   [<span data-ttu-id="eef7d-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eef7d-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="eef7d-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="eef7d-211">Requirements</span></span>

|<span data-ttu-id="eef7d-212">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-212">Requirement</span></span>| <span data-ttu-id="eef7d-213">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-215">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-215">1.1</span></span>|
|[<span data-ttu-id="eef7d-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="eef7d-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="eef7d-217">制限あり</span><span class="sxs-lookup"><span data-stu-id="eef7d-217">Restricted</span></span>|
|[<span data-ttu-id="eef7d-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-219">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="eef7d-220">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="eef7d-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="eef7d-221">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="eef7d-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="eef7d-222">種類</span><span class="sxs-lookup"><span data-stu-id="eef7d-222">Type</span></span>

*   [<span data-ttu-id="eef7d-223">UI</span><span class="sxs-lookup"><span data-stu-id="eef7d-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="eef7d-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="eef7d-224">Requirements</span></span>

|<span data-ttu-id="eef7d-225">要件</span><span class="sxs-lookup"><span data-stu-id="eef7d-225">Requirement</span></span>| <span data-ttu-id="eef7d-226">値</span><span class="sxs-lookup"><span data-stu-id="eef7d-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="eef7d-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="eef7d-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eef7d-228">1.1</span><span class="sxs-lookup"><span data-stu-id="eef7d-228">1.1</span></span>|
|[<span data-ttu-id="eef7d-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="eef7d-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eef7d-230">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="eef7d-230">Compose or Read</span></span>|
