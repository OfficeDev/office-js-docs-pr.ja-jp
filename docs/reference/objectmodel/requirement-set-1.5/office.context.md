---
title: Office.context - 要件セット 1.5
description: Outlook アドイン API の Outlook コンテキストオブジェクトのオブジェクトモデル (Mailbox API 1.5 バージョン)。
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0a226b796a3ac31729b08d68920a060094604a9f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717713"
---
# <a name="context"></a><span data-ttu-id="18822-103">context</span><span class="sxs-lookup"><span data-stu-id="18822-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="18822-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="18822-104">[Office](office.md).context</span></span>

<span data-ttu-id="18822-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="18822-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="18822-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.5)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="18822-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="18822-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-107">Requirements</span></span>

|<span data-ttu-id="18822-108">要件</span><span class="sxs-lookup"><span data-stu-id="18822-108">Requirement</span></span>| <span data-ttu-id="18822-109">値</span><span class="sxs-lookup"><span data-stu-id="18822-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-111">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-111">1.1</span></span>|
|[<span data-ttu-id="18822-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="18822-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="18822-114">Properties</span></span>

| <span data-ttu-id="18822-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="18822-115">Property</span></span> | <span data-ttu-id="18822-116">モード</span><span class="sxs-lookup"><span data-stu-id="18822-116">Modes</span></span> | <span data-ttu-id="18822-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="18822-117">Return type</span></span> | <span data-ttu-id="18822-118">最小値</span><span class="sxs-lookup"><span data-stu-id="18822-118">Minimum</span></span><br><span data-ttu-id="18822-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="18822-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="18822-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="18822-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="18822-121">作成</span><span class="sxs-lookup"><span data-stu-id="18822-121">Compose</span></span><br><span data-ttu-id="18822-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-122">Read</span></span> | <span data-ttu-id="18822-123">文字列</span><span class="sxs-lookup"><span data-stu-id="18822-123">String</span></span> | [<span data-ttu-id="18822-124">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-125">ダン</span><span class="sxs-lookup"><span data-stu-id="18822-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="18822-126">作成</span><span class="sxs-lookup"><span data-stu-id="18822-126">Compose</span></span><br><span data-ttu-id="18822-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-127">Read</span></span> | [<span data-ttu-id="18822-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="18822-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="18822-129">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="18822-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="18822-131">作成</span><span class="sxs-lookup"><span data-stu-id="18822-131">Compose</span></span><br><span data-ttu-id="18822-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-132">Read</span></span> | <span data-ttu-id="18822-133">文字列</span><span class="sxs-lookup"><span data-stu-id="18822-133">String</span></span> | [<span data-ttu-id="18822-134">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-135">主催</span><span class="sxs-lookup"><span data-stu-id="18822-135">host</span></span>](#host-hosttype) | <span data-ttu-id="18822-136">作成</span><span class="sxs-lookup"><span data-stu-id="18822-136">Compose</span></span><br><span data-ttu-id="18822-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-137">Read</span></span> | [<span data-ttu-id="18822-138">HostType</span><span class="sxs-lookup"><span data-stu-id="18822-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="18822-139">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="18822-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="18822-141">作成</span><span class="sxs-lookup"><span data-stu-id="18822-141">Compose</span></span><br><span data-ttu-id="18822-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-142">Read</span></span> | [<span data-ttu-id="18822-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="18822-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="18822-144">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="18822-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="18822-146">作成</span><span class="sxs-lookup"><span data-stu-id="18822-146">Compose</span></span><br><span data-ttu-id="18822-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-147">Read</span></span> | [<span data-ttu-id="18822-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="18822-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="18822-149">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-150">要件</span><span class="sxs-lookup"><span data-stu-id="18822-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="18822-151">作成</span><span class="sxs-lookup"><span data-stu-id="18822-151">Compose</span></span><br><span data-ttu-id="18822-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-152">Read</span></span> | [<span data-ttu-id="18822-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="18822-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="18822-154">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="18822-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="18822-156">作成</span><span class="sxs-lookup"><span data-stu-id="18822-156">Compose</span></span><br><span data-ttu-id="18822-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-157">Read</span></span> | [<span data-ttu-id="18822-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="18822-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="18822-159">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="18822-160">UI</span><span class="sxs-lookup"><span data-stu-id="18822-160">ui</span></span>](#ui-ui) | <span data-ttu-id="18822-161">作成</span><span class="sxs-lookup"><span data-stu-id="18822-161">Compose</span></span><br><span data-ttu-id="18822-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="18822-162">Read</span></span> | [<span data-ttu-id="18822-163">UI</span><span class="sxs-lookup"><span data-stu-id="18822-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="18822-164">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="18822-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="18822-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="18822-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="18822-166">contentLanguage: String</span></span>

<span data-ttu-id="18822-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="18822-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="18822-168">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="18822-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-169">型</span><span class="sxs-lookup"><span data-stu-id="18822-169">Type</span></span>

*   <span data-ttu-id="18822-170">String</span><span class="sxs-lookup"><span data-stu-id="18822-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18822-171">要件</span><span class="sxs-lookup"><span data-stu-id="18822-171">Requirements</span></span>

|<span data-ttu-id="18822-172">要件</span><span class="sxs-lookup"><span data-stu-id="18822-172">Requirement</span></span>| <span data-ttu-id="18822-173">値</span><span class="sxs-lookup"><span data-stu-id="18822-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-175">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-175">1.1</span></span>|
|[<span data-ttu-id="18822-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-178">例</span><span class="sxs-lookup"><span data-stu-id="18822-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="18822-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="18822-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="18822-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="18822-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-181">種類</span><span class="sxs-lookup"><span data-stu-id="18822-181">Type</span></span>

*   [<span data-ttu-id="18822-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="18822-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="18822-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-183">Requirements</span></span>

|<span data-ttu-id="18822-184">要件</span><span class="sxs-lookup"><span data-stu-id="18822-184">Requirement</span></span>| <span data-ttu-id="18822-185">値</span><span class="sxs-lookup"><span data-stu-id="18822-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-187">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-187">1.1</span></span>|
|[<span data-ttu-id="18822-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-190">例</span><span class="sxs-lookup"><span data-stu-id="18822-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="18822-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="18822-191">displayLanguage: String</span></span>

<span data-ttu-id="18822-192">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="18822-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="18822-193">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="18822-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-194">型</span><span class="sxs-lookup"><span data-stu-id="18822-194">Type</span></span>

*   <span data-ttu-id="18822-195">String</span><span class="sxs-lookup"><span data-stu-id="18822-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="18822-196">要件</span><span class="sxs-lookup"><span data-stu-id="18822-196">Requirements</span></span>

|<span data-ttu-id="18822-197">要件</span><span class="sxs-lookup"><span data-stu-id="18822-197">Requirement</span></span>| <span data-ttu-id="18822-198">値</span><span class="sxs-lookup"><span data-stu-id="18822-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-200">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-200">1.1</span></span>|
|[<span data-ttu-id="18822-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-203">例</span><span class="sxs-lookup"><span data-stu-id="18822-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="18822-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="18822-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="18822-205">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="18822-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-206">種類</span><span class="sxs-lookup"><span data-stu-id="18822-206">Type</span></span>

*   [<span data-ttu-id="18822-207">HostType</span><span class="sxs-lookup"><span data-stu-id="18822-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="18822-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-208">Requirements</span></span>

|<span data-ttu-id="18822-209">要件</span><span class="sxs-lookup"><span data-stu-id="18822-209">Requirement</span></span>| <span data-ttu-id="18822-210">値</span><span class="sxs-lookup"><span data-stu-id="18822-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-212">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-212">1.1</span></span>|
|[<span data-ttu-id="18822-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-215">例</span><span class="sxs-lookup"><span data-stu-id="18822-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="18822-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="18822-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="18822-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="18822-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-218">種類</span><span class="sxs-lookup"><span data-stu-id="18822-218">Type</span></span>

*   [<span data-ttu-id="18822-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="18822-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="18822-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-220">Requirements</span></span>

|<span data-ttu-id="18822-221">要件</span><span class="sxs-lookup"><span data-stu-id="18822-221">Requirement</span></span>| <span data-ttu-id="18822-222">値</span><span class="sxs-lookup"><span data-stu-id="18822-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-224">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-224">1.1</span></span>|
|[<span data-ttu-id="18822-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-227">例</span><span class="sxs-lookup"><span data-stu-id="18822-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="18822-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="18822-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="18822-229">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="18822-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-230">種類</span><span class="sxs-lookup"><span data-stu-id="18822-230">Type</span></span>

*   [<span data-ttu-id="18822-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="18822-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="18822-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-232">Requirements</span></span>

|<span data-ttu-id="18822-233">要件</span><span class="sxs-lookup"><span data-stu-id="18822-233">Requirement</span></span>| <span data-ttu-id="18822-234">値</span><span class="sxs-lookup"><span data-stu-id="18822-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-236">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-236">1.1</span></span>|
|[<span data-ttu-id="18822-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="18822-239">例</span><span class="sxs-lookup"><span data-stu-id="18822-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="18822-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="18822-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="18822-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="18822-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="18822-242">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="18822-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-243">種類</span><span class="sxs-lookup"><span data-stu-id="18822-243">Type</span></span>

*   [<span data-ttu-id="18822-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="18822-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="18822-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-245">Requirements</span></span>

|<span data-ttu-id="18822-246">要件</span><span class="sxs-lookup"><span data-stu-id="18822-246">Requirement</span></span>| <span data-ttu-id="18822-247">値</span><span class="sxs-lookup"><span data-stu-id="18822-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-249">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-249">1.1</span></span>|
|[<span data-ttu-id="18822-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="18822-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="18822-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="18822-251">Restricted</span></span>|
|[<span data-ttu-id="18822-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="18822-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="18822-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="18822-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="18822-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="18822-256">種類</span><span class="sxs-lookup"><span data-stu-id="18822-256">Type</span></span>

*   [<span data-ttu-id="18822-257">UI</span><span class="sxs-lookup"><span data-stu-id="18822-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="18822-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="18822-258">Requirements</span></span>

|<span data-ttu-id="18822-259">要件</span><span class="sxs-lookup"><span data-stu-id="18822-259">Requirement</span></span>| <span data-ttu-id="18822-260">値</span><span class="sxs-lookup"><span data-stu-id="18822-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="18822-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="18822-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="18822-262">1.1</span><span class="sxs-lookup"><span data-stu-id="18822-262">1.1</span></span>|
|[<span data-ttu-id="18822-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="18822-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="18822-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="18822-264">Compose or Read</span></span>|
