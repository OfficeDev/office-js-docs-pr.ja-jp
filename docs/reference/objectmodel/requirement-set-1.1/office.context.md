---
title: Office.context - 要件セット 1.1
description: Office。メールボックス API 要件セット 1.1 をOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 41273bfc5362a9d5572e38b8e80b81041f5aa312
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590877"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="0ab34-103">context (メールボックス要件セット 1.1)</span><span class="sxs-lookup"><span data-stu-id="0ab34-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="0ab34-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="0ab34-104">[Office](office.md).context</span></span>

<span data-ttu-id="0ab34-105">Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="0ab34-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="0ab34-106">この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="0ab34-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ab34-107">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-107">Requirements</span></span>

|<span data-ttu-id="0ab34-108">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-108">Requirement</span></span>| <span data-ttu-id="0ab34-109">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-111">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-111">1.1</span></span>|
|[<span data-ttu-id="0ab34-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="0ab34-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0ab34-114">Properties</span></span>

| <span data-ttu-id="0ab34-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="0ab34-115">Property</span></span> | <span data-ttu-id="0ab34-116">モード</span><span class="sxs-lookup"><span data-stu-id="0ab34-116">Modes</span></span> | <span data-ttu-id="0ab34-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="0ab34-117">Return type</span></span> | <span data-ttu-id="0ab34-118">最小値</span><span class="sxs-lookup"><span data-stu-id="0ab34-118">Minimum</span></span><br><span data-ttu-id="0ab34-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="0ab34-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0ab34-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="0ab34-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="0ab34-121">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-121">Compose</span></span><br><span data-ttu-id="0ab34-122">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-122">Read</span></span> | <span data-ttu-id="0ab34-123">String</span><span class="sxs-lookup"><span data-stu-id="0ab34-123">String</span></span> | [<span data-ttu-id="0ab34-124">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-125">診断</span><span class="sxs-lookup"><span data-stu-id="0ab34-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="0ab34-126">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-126">Compose</span></span><br><span data-ttu-id="0ab34-127">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-127">Read</span></span> | [<span data-ttu-id="0ab34-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="0ab34-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="0ab34-129">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="0ab34-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="0ab34-131">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-131">Compose</span></span><br><span data-ttu-id="0ab34-132">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-132">Read</span></span> | <span data-ttu-id="0ab34-133">String</span><span class="sxs-lookup"><span data-stu-id="0ab34-133">String</span></span> | [<span data-ttu-id="0ab34-134">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="0ab34-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="0ab34-136">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-136">Compose</span></span><br><span data-ttu-id="0ab34-137">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-137">Read</span></span> | [<span data-ttu-id="0ab34-138">メールボックス</span><span class="sxs-lookup"><span data-stu-id="0ab34-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="0ab34-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-140">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="0ab34-141">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-141">Compose</span></span><br><span data-ttu-id="0ab34-142">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-142">Read</span></span> | [<span data-ttu-id="0ab34-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="0ab34-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="0ab34-144">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="0ab34-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="0ab34-146">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-146">Compose</span></span><br><span data-ttu-id="0ab34-147">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-147">Read</span></span> | [<span data-ttu-id="0ab34-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0ab34-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="0ab34-149">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0ab34-150">UI</span><span class="sxs-lookup"><span data-stu-id="0ab34-150">ui</span></span>](#ui-ui) | <span data-ttu-id="0ab34-151">作成</span><span class="sxs-lookup"><span data-stu-id="0ab34-151">Compose</span></span><br><span data-ttu-id="0ab34-152">Read</span><span class="sxs-lookup"><span data-stu-id="0ab34-152">Read</span></span> | [<span data-ttu-id="0ab34-153">UI</span><span class="sxs-lookup"><span data-stu-id="0ab34-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="0ab34-154">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="0ab34-155">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="0ab34-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="0ab34-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="0ab34-156">contentLanguage: String</span></span>

<span data-ttu-id="0ab34-157">アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ab34-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="0ab34-158">この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 </span><span class="sxs-lookup"><span data-stu-id="0ab34-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-159">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-159">Type</span></span>

*   <span data-ttu-id="0ab34-160">String</span><span class="sxs-lookup"><span data-stu-id="0ab34-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ab34-161">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-161">Requirements</span></span>

|<span data-ttu-id="0ab34-162">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-162">Requirement</span></span>| <span data-ttu-id="0ab34-163">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-164">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-165">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-165">1.1</span></span>|
|[<span data-ttu-id="0ab34-166">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-167">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ab34-168">例</span><span class="sxs-lookup"><span data-stu-id="0ab34-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="0ab34-169">診断: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="0ab34-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="0ab34-170">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="0ab34-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-171">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-171">Type</span></span>

*   [<span data-ttu-id="0ab34-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="0ab34-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="0ab34-173">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-173">Requirements</span></span>

|<span data-ttu-id="0ab34-174">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-174">Requirement</span></span>| <span data-ttu-id="0ab34-175">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-176">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-177">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-177">1.1</span></span>|
|[<span data-ttu-id="0ab34-178">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-179">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ab34-180">例</span><span class="sxs-lookup"><span data-stu-id="0ab34-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="0ab34-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="0ab34-181">displayLanguage: String</span></span>

<span data-ttu-id="0ab34-182">ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="0ab34-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="0ab34-183">この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。</span><span class="sxs-lookup"><span data-stu-id="0ab34-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-184">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-184">Type</span></span>

*   <span data-ttu-id="0ab34-185">String</span><span class="sxs-lookup"><span data-stu-id="0ab34-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0ab34-186">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-186">Requirements</span></span>

|<span data-ttu-id="0ab34-187">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-187">Requirement</span></span>| <span data-ttu-id="0ab34-188">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-189">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-190">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-190">1.1</span></span>|
|[<span data-ttu-id="0ab34-191">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-192">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ab34-193">例</span><span class="sxs-lookup"><span data-stu-id="0ab34-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="0ab34-194">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="0ab34-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="0ab34-195">現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="0ab34-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-196">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-196">Type</span></span>

*   [<span data-ttu-id="0ab34-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="0ab34-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="0ab34-198">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-198">Requirements</span></span>

|<span data-ttu-id="0ab34-199">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-199">Requirement</span></span>| <span data-ttu-id="0ab34-200">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-201">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-202">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-202">1.1</span></span>|
|[<span data-ttu-id="0ab34-203">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-204">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0ab34-205">例</span><span class="sxs-lookup"><span data-stu-id="0ab34-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="0ab34-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="0ab34-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="0ab34-207">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="0ab34-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="0ab34-208">このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`</span><span class="sxs-lookup"><span data-stu-id="0ab34-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-209">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-209">Type</span></span>

*   [<span data-ttu-id="0ab34-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0ab34-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="0ab34-211">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-211">Requirements</span></span>

|<span data-ttu-id="0ab34-212">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-212">Requirement</span></span>| <span data-ttu-id="0ab34-213">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-214">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-215">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-215">1.1</span></span>|
|[<span data-ttu-id="0ab34-216">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="0ab34-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="0ab34-217">制限あり</span><span class="sxs-lookup"><span data-stu-id="0ab34-217">Restricted</span></span>|
|[<span data-ttu-id="0ab34-218">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-219">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="0ab34-220">ui: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="0ab34-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="0ab34-221">ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。</span><span class="sxs-lookup"><span data-stu-id="0ab34-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0ab34-222">型</span><span class="sxs-lookup"><span data-stu-id="0ab34-222">Type</span></span>

*   [<span data-ttu-id="0ab34-223">UI</span><span class="sxs-lookup"><span data-stu-id="0ab34-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="0ab34-224">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-224">Requirements</span></span>

|<span data-ttu-id="0ab34-225">要件</span><span class="sxs-lookup"><span data-stu-id="0ab34-225">Requirement</span></span>| <span data-ttu-id="0ab34-226">値</span><span class="sxs-lookup"><span data-stu-id="0ab34-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="0ab34-227">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="0ab34-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0ab34-228">1.1</span><span class="sxs-lookup"><span data-stu-id="0ab34-228">1.1</span></span>|
|[<span data-ttu-id="0ab34-229">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="0ab34-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0ab34-230">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="0ab34-230">Compose or Read</span></span>|
