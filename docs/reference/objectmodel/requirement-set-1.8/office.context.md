---
title: Office.context - 要件セット 1.8
description: Office。メールボックス API 要件セット 1.8 をOutlookアドインで使用できるコンテキスト オブジェクト メンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 99573d9984c571c99461e90e8bdccdca35fe30b7
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590968"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="7017e-103">context (メールボックス要件セット 1.8)</span><span class="sxs-lookup"><span data-stu-id="7017e-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="7017e-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="7017e-104">[Office](office.md).context</span></span>

<span data-ttu-id="7017e-105">Office.context は、すべてのアプリでアドインによって使用される共有インターフェイスをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="7017e-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="7017e-106">この一覧には、アドインで使用されるインターフェイスOutlook記載されています。Office.context 名前空間の完全な一覧については、common API の[Office.context リファレンスを参照してください](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true)。</span><span class="sxs-lookup"><span data-stu-id="7017e-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7017e-107">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-107">Requirements</span></span>

|<span data-ttu-id="7017e-108">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-108">Requirement</span></span>| <span data-ttu-id="7017e-109">値</span><span class="sxs-lookup"><span data-stu-id="7017e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-111">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-111">1.1</span></span>|
|[<span data-ttu-id="7017e-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="7017e-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7017e-114">Properties</span></span>

| <span data-ttu-id="7017e-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="7017e-115">Property</span></span> | <span data-ttu-id="7017e-116">モード</span><span class="sxs-lookup"><span data-stu-id="7017e-116">Modes</span></span> | <span data-ttu-id="7017e-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="7017e-117">Return type</span></span> | <span data-ttu-id="7017e-118">最小値</span><span class="sxs-lookup"><span data-stu-id="7017e-118">Minimum</span></span><br><span data-ttu-id="7017e-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="7017e-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7017e-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="7017e-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="7017e-121">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-121">Compose</span></span><br><span data-ttu-id="7017e-122">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-122">Read</span></span> | <span data-ttu-id="7017e-123">String</span><span class="sxs-lookup"><span data-stu-id="7017e-123">String</span></span> | [<span data-ttu-id="7017e-124">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-125">診断</span><span class="sxs-lookup"><span data-stu-id="7017e-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="7017e-126">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-126">Compose</span></span><br><span data-ttu-id="7017e-127">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-127">Read</span></span> | [<span data-ttu-id="7017e-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7017e-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7017e-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7017e-131">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-131">Compose</span></span><br><span data-ttu-id="7017e-132">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-132">Read</span></span> | <span data-ttu-id="7017e-133">String</span><span class="sxs-lookup"><span data-stu-id="7017e-133">String</span></span> | [<span data-ttu-id="7017e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-135">host</span><span class="sxs-lookup"><span data-stu-id="7017e-135">host</span></span>](#host-hosttype) | <span data-ttu-id="7017e-136">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-136">Compose</span></span><br><span data-ttu-id="7017e-137">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-137">Read</span></span> | [<span data-ttu-id="7017e-138">HostType</span><span class="sxs-lookup"><span data-stu-id="7017e-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-139">1.5</span><span class="sxs-lookup"><span data-stu-id="7017e-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7017e-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="7017e-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="7017e-141">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-141">Compose</span></span><br><span data-ttu-id="7017e-142">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-142">Read</span></span> | [<span data-ttu-id="7017e-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="7017e-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-144">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7017e-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="7017e-146">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-146">Compose</span></span><br><span data-ttu-id="7017e-147">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-147">Read</span></span> | [<span data-ttu-id="7017e-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7017e-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-149">1.5</span><span class="sxs-lookup"><span data-stu-id="7017e-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7017e-150">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="7017e-151">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-151">Compose</span></span><br><span data-ttu-id="7017e-152">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-152">Read</span></span> | [<span data-ttu-id="7017e-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7017e-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-154">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7017e-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7017e-156">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-156">Compose</span></span><br><span data-ttu-id="7017e-157">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-157">Read</span></span> | [<span data-ttu-id="7017e-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7017e-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7017e-160">UI</span><span class="sxs-lookup"><span data-stu-id="7017e-160">ui</span></span>](#ui-ui) | <span data-ttu-id="7017e-161">作成</span><span class="sxs-lookup"><span data-stu-id="7017e-161">Compose</span></span><br><span data-ttu-id="7017e-162">Read</span><span class="sxs-lookup"><span data-stu-id="7017e-162">Read</span></span> | [<span data-ttu-id="7017e-163">UI</span><span class="sxs-lookup"><span data-stu-id="7017e-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="7017e-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="7017e-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="7017e-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="7017e-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7017e-166">contentLanguage: String</span></span>

<span data-ttu-id="7017e-167">アイテムを編集するユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="7017e-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="7017e-168">この値は、クライアント アプリケーション内の [ファイル] > オプション > `contentLanguage` **言語** でOffice設定を反映します。 </span><span class="sxs-lookup"><span data-stu-id="7017e-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-169">型</span><span class="sxs-lookup"><span data-stu-id="7017e-169">Type</span></span>

*   <span data-ttu-id="7017e-170">String</span><span class="sxs-lookup"><span data-stu-id="7017e-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7017e-171">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-171">Requirements</span></span>

|<span data-ttu-id="7017e-172">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-172">Requirement</span></span>| <span data-ttu-id="7017e-173">値</span><span class="sxs-lookup"><span data-stu-id="7017e-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-175">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-175">1.1</span></span>|
|[<span data-ttu-id="7017e-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-178">例</span><span class="sxs-lookup"><span data-stu-id="7017e-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="7017e-179">診断: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="7017e-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="7017e-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="7017e-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-181">型</span><span class="sxs-lookup"><span data-stu-id="7017e-181">Type</span></span>

*   [<span data-ttu-id="7017e-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7017e-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="7017e-183">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-183">Requirements</span></span>

|<span data-ttu-id="7017e-184">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-184">Requirement</span></span>| <span data-ttu-id="7017e-185">値</span><span class="sxs-lookup"><span data-stu-id="7017e-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-187">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-187">1.1</span></span>|
|[<span data-ttu-id="7017e-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-190">例</span><span class="sxs-lookup"><span data-stu-id="7017e-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="7017e-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7017e-191">displayLanguage: String</span></span>

<span data-ttu-id="7017e-192">ユーザーがクライアント アプリケーションの UI 用に指定した RFC 1766 Language タグ形式のロケール (言語) をOfficeします。</span><span class="sxs-lookup"><span data-stu-id="7017e-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="7017e-193">この `displayLanguage` 値は、クライアントアプリケーションの [File >**オプション**] >言語でOffice反映されます。</span><span class="sxs-lookup"><span data-stu-id="7017e-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-194">型</span><span class="sxs-lookup"><span data-stu-id="7017e-194">Type</span></span>

*   <span data-ttu-id="7017e-195">String</span><span class="sxs-lookup"><span data-stu-id="7017e-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7017e-196">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-196">Requirements</span></span>

|<span data-ttu-id="7017e-197">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-197">Requirement</span></span>| <span data-ttu-id="7017e-198">値</span><span class="sxs-lookup"><span data-stu-id="7017e-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-200">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-200">1.1</span></span>|
|[<span data-ttu-id="7017e-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-203">例</span><span class="sxs-lookup"><span data-stu-id="7017e-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="7017e-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="7017e-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="7017e-205">アドインをOfficeしているアプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="7017e-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7017e-206">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してホストを取得できます。</span><span class="sxs-lookup"><span data-stu-id="7017e-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-207">型</span><span class="sxs-lookup"><span data-stu-id="7017e-207">Type</span></span>

*   [<span data-ttu-id="7017e-208">HostType</span><span class="sxs-lookup"><span data-stu-id="7017e-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="7017e-209">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-209">Requirements</span></span>

|<span data-ttu-id="7017e-210">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-210">Requirement</span></span>| <span data-ttu-id="7017e-211">値</span><span class="sxs-lookup"><span data-stu-id="7017e-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-213">1.5</span><span class="sxs-lookup"><span data-stu-id="7017e-213">1.5</span></span>|
|[<span data-ttu-id="7017e-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-215">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-216">例</span><span class="sxs-lookup"><span data-stu-id="7017e-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="7017e-217">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="7017e-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="7017e-218">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="7017e-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="7017e-219">または[、Office.context.diagnostics](#diagnostics-contextinformation)プロパティを使用してプラットフォームを取得できます。</span><span class="sxs-lookup"><span data-stu-id="7017e-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-220">型</span><span class="sxs-lookup"><span data-stu-id="7017e-220">Type</span></span>

*   [<span data-ttu-id="7017e-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7017e-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="7017e-222">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-222">Requirements</span></span>

|<span data-ttu-id="7017e-223">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-223">Requirement</span></span>| <span data-ttu-id="7017e-224">値</span><span class="sxs-lookup"><span data-stu-id="7017e-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-226">1.5</span><span class="sxs-lookup"><span data-stu-id="7017e-226">1.5</span></span>|
|[<span data-ttu-id="7017e-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-229">例</span><span class="sxs-lookup"><span data-stu-id="7017e-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="7017e-230">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="7017e-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="7017e-231">現在のアプリケーションとプラットフォームでサポートされている要件セットを決定するメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="7017e-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-232">型</span><span class="sxs-lookup"><span data-stu-id="7017e-232">Type</span></span>

*   [<span data-ttu-id="7017e-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7017e-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="7017e-234">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-234">Requirements</span></span>

|<span data-ttu-id="7017e-235">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-235">Requirement</span></span>| <span data-ttu-id="7017e-236">値</span><span class="sxs-lookup"><span data-stu-id="7017e-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-238">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-238">1.1</span></span>|
|[<span data-ttu-id="7017e-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-240">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7017e-241">例</span><span class="sxs-lookup"><span data-stu-id="7017e-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="7017e-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="7017e-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="7017e-243">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="7017e-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7017e-244">このオブジェクトを使用すると、ユーザーのメールボックスに格納されているメール アドインのデータを格納してアクセスできます。これにより、そのメールボックスへのアクセスに使用される Outlook クライアントから実行されている場合に、そのアドインが使用できます。 `RoamingSettings`</span><span class="sxs-lookup"><span data-stu-id="7017e-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-245">型</span><span class="sxs-lookup"><span data-stu-id="7017e-245">Type</span></span>

*   [<span data-ttu-id="7017e-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7017e-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7017e-247">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-247">Requirements</span></span>

|<span data-ttu-id="7017e-248">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-248">Requirement</span></span>| <span data-ttu-id="7017e-249">値</span><span class="sxs-lookup"><span data-stu-id="7017e-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-251">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-251">1.1</span></span>|
|[<span data-ttu-id="7017e-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7017e-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="7017e-253">制限あり</span><span class="sxs-lookup"><span data-stu-id="7017e-253">Restricted</span></span>|
|[<span data-ttu-id="7017e-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="7017e-256">ui: [UI](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="7017e-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="7017e-257">ダイアログ ボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを、Office提供します。</span><span class="sxs-lookup"><span data-stu-id="7017e-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7017e-258">型</span><span class="sxs-lookup"><span data-stu-id="7017e-258">Type</span></span>

*   [<span data-ttu-id="7017e-259">UI</span><span class="sxs-lookup"><span data-stu-id="7017e-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="7017e-260">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-260">Requirements</span></span>

|<span data-ttu-id="7017e-261">要件</span><span class="sxs-lookup"><span data-stu-id="7017e-261">Requirement</span></span>| <span data-ttu-id="7017e-262">値</span><span class="sxs-lookup"><span data-stu-id="7017e-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="7017e-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7017e-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7017e-264">1.1</span><span class="sxs-lookup"><span data-stu-id="7017e-264">1.1</span></span>|
|[<span data-ttu-id="7017e-265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7017e-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7017e-266">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7017e-266">Compose or Read</span></span>|
