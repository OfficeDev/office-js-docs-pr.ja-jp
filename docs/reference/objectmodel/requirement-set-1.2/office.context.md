---
title: Office コンテキスト要件セット1.2
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 2c6a4709f5771b4bdb9cdc40b028770fbfc3a63e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165448"
---
# <a name="context"></a><span data-ttu-id="1592e-102">context</span><span class="sxs-lookup"><span data-stu-id="1592e-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="1592e-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1592e-103">[Office](office.md).context</span></span>

<span data-ttu-id="1592e-104">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1592e-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1592e-105">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.2)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1592e-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1592e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-106">Requirements</span></span>

|<span data-ttu-id="1592e-107">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-107">Requirement</span></span>| <span data-ttu-id="1592e-108">値</span><span class="sxs-lookup"><span data-stu-id="1592e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-110">1.1</span></span>|
|[<span data-ttu-id="1592e-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1592e-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1592e-113">Properties</span></span>

| <span data-ttu-id="1592e-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="1592e-114">Property</span></span> | <span data-ttu-id="1592e-115">モード</span><span class="sxs-lookup"><span data-stu-id="1592e-115">Modes</span></span> | <span data-ttu-id="1592e-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="1592e-116">Return type</span></span> | <span data-ttu-id="1592e-117">最小値</span><span class="sxs-lookup"><span data-stu-id="1592e-117">Minimum</span></span><br><span data-ttu-id="1592e-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="1592e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1592e-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1592e-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1592e-120">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-120">Compose</span></span><br><span data-ttu-id="1592e-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-121">Read</span></span> | <span data-ttu-id="1592e-122">文字列</span><span class="sxs-lookup"><span data-stu-id="1592e-122">String</span></span> | [<span data-ttu-id="1592e-123">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-124">ダン</span><span class="sxs-lookup"><span data-stu-id="1592e-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1592e-125">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-125">Compose</span></span><br><span data-ttu-id="1592e-126">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-126">Read</span></span> | [<span data-ttu-id="1592e-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1592e-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [<span data-ttu-id="1592e-128">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1592e-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1592e-130">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-130">Compose</span></span><br><span data-ttu-id="1592e-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-131">Read</span></span> | <span data-ttu-id="1592e-132">文字列</span><span class="sxs-lookup"><span data-stu-id="1592e-132">String</span></span> | [<span data-ttu-id="1592e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-134">主催</span><span class="sxs-lookup"><span data-stu-id="1592e-134">host</span></span>](#host-hosttype) | <span data-ttu-id="1592e-135">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-135">Compose</span></span><br><span data-ttu-id="1592e-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-136">Read</span></span> | [<span data-ttu-id="1592e-137">HostType</span><span class="sxs-lookup"><span data-stu-id="1592e-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [<span data-ttu-id="1592e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="1592e-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1592e-140">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-140">Compose</span></span><br><span data-ttu-id="1592e-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-141">Read</span></span> | [<span data-ttu-id="1592e-142">メールボックス</span><span class="sxs-lookup"><span data-stu-id="1592e-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2) | [<span data-ttu-id="1592e-143">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-144">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1592e-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1592e-145">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-145">Compose</span></span><br><span data-ttu-id="1592e-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-146">Read</span></span> | [<span data-ttu-id="1592e-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1592e-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [<span data-ttu-id="1592e-148">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-149">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1592e-150">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-150">Compose</span></span><br><span data-ttu-id="1592e-151">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-151">Read</span></span> | [<span data-ttu-id="1592e-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1592e-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [<span data-ttu-id="1592e-153">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1592e-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1592e-155">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-155">Compose</span></span><br><span data-ttu-id="1592e-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-156">Read</span></span> | [<span data-ttu-id="1592e-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1592e-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [<span data-ttu-id="1592e-158">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1592e-159">UI</span><span class="sxs-lookup"><span data-stu-id="1592e-159">ui</span></span>](#ui-ui) | <span data-ttu-id="1592e-160">作成</span><span class="sxs-lookup"><span data-stu-id="1592e-160">Compose</span></span><br><span data-ttu-id="1592e-161">読み取り</span><span class="sxs-lookup"><span data-stu-id="1592e-161">Read</span></span> | [<span data-ttu-id="1592e-162">UI</span><span class="sxs-lookup"><span data-stu-id="1592e-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2) | [<span data-ttu-id="1592e-163">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1592e-164">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="1592e-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1592e-165">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1592e-165">contentLanguage: String</span></span>

<span data-ttu-id="1592e-166">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="1592e-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1592e-167">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="1592e-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-168">型</span><span class="sxs-lookup"><span data-stu-id="1592e-168">Type</span></span>

*   <span data-ttu-id="1592e-169">String</span><span class="sxs-lookup"><span data-stu-id="1592e-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1592e-170">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-170">Requirements</span></span>

|<span data-ttu-id="1592e-171">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-171">Requirement</span></span>| <span data-ttu-id="1592e-172">値</span><span class="sxs-lookup"><span data-stu-id="1592e-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-174">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-174">1.1</span></span>|
|[<span data-ttu-id="1592e-175">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-175">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-176">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-177">例</span><span class="sxs-lookup"><span data-stu-id="1592e-177">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1592e-178">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1592e-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1592e-179">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="1592e-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-180">型</span><span class="sxs-lookup"><span data-stu-id="1592e-180">Type</span></span>

*   [<span data-ttu-id="1592e-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1592e-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1592e-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-182">Requirements</span></span>

|<span data-ttu-id="1592e-183">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-183">Requirement</span></span>| <span data-ttu-id="1592e-184">値</span><span class="sxs-lookup"><span data-stu-id="1592e-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-186">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-186">1.1</span></span>|
|[<span data-ttu-id="1592e-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-187">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-188">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-189">例</span><span class="sxs-lookup"><span data-stu-id="1592e-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1592e-190">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1592e-190">displayLanguage: String</span></span>

<span data-ttu-id="1592e-191">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="1592e-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1592e-192">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="1592e-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-193">型</span><span class="sxs-lookup"><span data-stu-id="1592e-193">Type</span></span>

*   <span data-ttu-id="1592e-194">String</span><span class="sxs-lookup"><span data-stu-id="1592e-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1592e-195">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-195">Requirements</span></span>

|<span data-ttu-id="1592e-196">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-196">Requirement</span></span>| <span data-ttu-id="1592e-197">値</span><span class="sxs-lookup"><span data-stu-id="1592e-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-199">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-199">1.1</span></span>|
|[<span data-ttu-id="1592e-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-200">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-202">例</span><span class="sxs-lookup"><span data-stu-id="1592e-202">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1592e-203">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1592e-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1592e-204">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="1592e-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-205">型</span><span class="sxs-lookup"><span data-stu-id="1592e-205">Type</span></span>

*   [<span data-ttu-id="1592e-206">HostType</span><span class="sxs-lookup"><span data-stu-id="1592e-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1592e-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-207">Requirements</span></span>

|<span data-ttu-id="1592e-208">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-208">Requirement</span></span>| <span data-ttu-id="1592e-209">値</span><span class="sxs-lookup"><span data-stu-id="1592e-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-211">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-211">1.1</span></span>|
|[<span data-ttu-id="1592e-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-212">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-213">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-214">例</span><span class="sxs-lookup"><span data-stu-id="1592e-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1592e-215">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1592e-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1592e-216">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="1592e-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-217">型</span><span class="sxs-lookup"><span data-stu-id="1592e-217">Type</span></span>

*   [<span data-ttu-id="1592e-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1592e-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1592e-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-219">Requirements</span></span>

|<span data-ttu-id="1592e-220">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-220">Requirement</span></span>| <span data-ttu-id="1592e-221">値</span><span class="sxs-lookup"><span data-stu-id="1592e-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-223">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-223">1.1</span></span>|
|[<span data-ttu-id="1592e-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-226">例</span><span class="sxs-lookup"><span data-stu-id="1592e-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1592e-227">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1592e-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1592e-228">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="1592e-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-229">型</span><span class="sxs-lookup"><span data-stu-id="1592e-229">Type</span></span>

*   [<span data-ttu-id="1592e-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1592e-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1592e-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-231">Requirements</span></span>

|<span data-ttu-id="1592e-232">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-232">Requirement</span></span>| <span data-ttu-id="1592e-233">値</span><span class="sxs-lookup"><span data-stu-id="1592e-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-235">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-235">1.1</span></span>|
|[<span data-ttu-id="1592e-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1592e-238">例</span><span class="sxs-lookup"><span data-stu-id="1592e-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1592e-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1592e-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1592e-240">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1592e-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1592e-241">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="1592e-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-242">型</span><span class="sxs-lookup"><span data-stu-id="1592e-242">Type</span></span>

*   [<span data-ttu-id="1592e-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1592e-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1592e-244">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-244">Requirements</span></span>

|<span data-ttu-id="1592e-245">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-245">Requirement</span></span>| <span data-ttu-id="1592e-246">値</span><span class="sxs-lookup"><span data-stu-id="1592e-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-247">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-248">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-248">1.1</span></span>|
|[<span data-ttu-id="1592e-249">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1592e-249">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1592e-250">制限あり</span><span class="sxs-lookup"><span data-stu-id="1592e-250">Restricted</span></span>|
|[<span data-ttu-id="1592e-251">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-251">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-252">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1592e-253">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1592e-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1592e-254">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="1592e-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1592e-255">型</span><span class="sxs-lookup"><span data-stu-id="1592e-255">Type</span></span>

*   [<span data-ttu-id="1592e-256">UI</span><span class="sxs-lookup"><span data-stu-id="1592e-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1592e-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="1592e-257">Requirements</span></span>

|<span data-ttu-id="1592e-258">要件</span><span class="sxs-lookup"><span data-stu-id="1592e-258">Requirement</span></span>| <span data-ttu-id="1592e-259">値</span><span class="sxs-lookup"><span data-stu-id="1592e-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="1592e-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1592e-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1592e-261">1.1</span><span class="sxs-lookup"><span data-stu-id="1592e-261">1.1</span></span>|
|[<span data-ttu-id="1592e-262">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1592e-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1592e-263">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1592e-263">Compose or Read</span></span>|
