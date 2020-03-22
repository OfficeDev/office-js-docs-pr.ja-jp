---
title: Office コンテキスト要件セット1.3
description: メールボックス API 要件セット1.3 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f6bc2c73ac40003910ae64ba7d58d8326b8e5bad
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890712"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="cb2e9-103">コンテキスト (メールボックス要件セット 1.3)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="cb2e9-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="cb2e9-104">[Office](office.md).context</span></span>

<span data-ttu-id="cb2e9-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="cb2e9-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.3)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb2e9-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-107">Requirements</span></span>

|<span data-ttu-id="cb2e9-108">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-108">Requirement</span></span>| <span data-ttu-id="cb2e9-109">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-111">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-111">1.1</span></span>|
|[<span data-ttu-id="cb2e9-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cb2e9-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="cb2e9-114">Properties</span></span>

| <span data-ttu-id="cb2e9-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="cb2e9-115">Property</span></span> | <span data-ttu-id="cb2e9-116">モード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-116">Modes</span></span> | <span data-ttu-id="cb2e9-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="cb2e9-117">Return type</span></span> | <span data-ttu-id="cb2e9-118">最小値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-118">Minimum</span></span><br><span data-ttu-id="cb2e9-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="cb2e9-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cb2e9-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="cb2e9-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="cb2e9-121">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-121">Compose</span></span><br><span data-ttu-id="cb2e9-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-122">Read</span></span> | <span data-ttu-id="cb2e9-123">String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-123">String</span></span> | [<span data-ttu-id="cb2e9-124">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-125">ダン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="cb2e9-126">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-126">Compose</span></span><br><span data-ttu-id="cb2e9-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-127">Read</span></span> | [<span data-ttu-id="cb2e9-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cb2e9-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-129">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="cb2e9-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="cb2e9-131">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-131">Compose</span></span><br><span data-ttu-id="cb2e9-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-132">Read</span></span> | <span data-ttu-id="cb2e9-133">String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-133">String</span></span> | [<span data-ttu-id="cb2e9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-135">主催</span><span class="sxs-lookup"><span data-stu-id="cb2e9-135">host</span></span>](#host-hosttype) | <span data-ttu-id="cb2e9-136">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-136">Compose</span></span><br><span data-ttu-id="cb2e9-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-137">Read</span></span> | [<span data-ttu-id="cb2e9-138">HostType</span><span class="sxs-lookup"><span data-stu-id="cb2e9-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="cb2e9-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="cb2e9-141">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-141">Compose</span></span><br><span data-ttu-id="cb2e9-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-142">Read</span></span> | [<span data-ttu-id="cb2e9-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="cb2e9-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-144">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="cb2e9-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="cb2e9-146">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-146">Compose</span></span><br><span data-ttu-id="cb2e9-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-147">Read</span></span> | [<span data-ttu-id="cb2e9-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="cb2e9-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-149">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-150">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="cb2e9-151">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-151">Compose</span></span><br><span data-ttu-id="cb2e9-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-152">Read</span></span> | [<span data-ttu-id="cb2e9-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cb2e9-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-154">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="cb2e9-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="cb2e9-156">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-156">Compose</span></span><br><span data-ttu-id="cb2e9-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-157">Read</span></span> | [<span data-ttu-id="cb2e9-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cb2e9-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-159">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cb2e9-160">UI</span><span class="sxs-lookup"><span data-stu-id="cb2e9-160">ui</span></span>](#ui-ui) | <span data-ttu-id="cb2e9-161">作成</span><span class="sxs-lookup"><span data-stu-id="cb2e9-161">Compose</span></span><br><span data-ttu-id="cb2e9-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="cb2e9-162">Read</span></span> | [<span data-ttu-id="cb2e9-163">UI</span><span class="sxs-lookup"><span data-stu-id="cb2e9-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="cb2e9-164">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="cb2e9-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="cb2e9-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="cb2e9-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-166">contentLanguage: String</span></span>

<span data-ttu-id="cb2e9-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="cb2e9-168">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-169">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-169">Type</span></span>

*   <span data-ttu-id="cb2e9-170">String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb2e9-171">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-171">Requirements</span></span>

|<span data-ttu-id="cb2e9-172">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-172">Requirement</span></span>| <span data-ttu-id="cb2e9-173">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-175">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-175">1.1</span></span>|
|[<span data-ttu-id="cb2e9-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-178">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="cb2e9-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="cb2e9-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-181">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-181">Type</span></span>

*   [<span data-ttu-id="cb2e9-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cb2e9-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-183">Requirements</span></span>

|<span data-ttu-id="cb2e9-184">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-184">Requirement</span></span>| <span data-ttu-id="cb2e9-185">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-187">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-187">1.1</span></span>|
|[<span data-ttu-id="cb2e9-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-190">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="cb2e9-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-191">displayLanguage: String</span></span>

<span data-ttu-id="cb2e9-192">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="cb2e9-193">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-194">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-194">Type</span></span>

*   <span data-ttu-id="cb2e9-195">String</span><span class="sxs-lookup"><span data-stu-id="cb2e9-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb2e9-196">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-196">Requirements</span></span>

|<span data-ttu-id="cb2e9-197">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-197">Requirement</span></span>| <span data-ttu-id="cb2e9-198">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-200">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-200">1.1</span></span>|
|[<span data-ttu-id="cb2e9-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-203">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="cb2e9-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="cb2e9-205">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-206">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-206">Type</span></span>

*   [<span data-ttu-id="cb2e9-207">HostType</span><span class="sxs-lookup"><span data-stu-id="cb2e9-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-208">Requirements</span></span>

|<span data-ttu-id="cb2e9-209">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-209">Requirement</span></span>| <span data-ttu-id="cb2e9-210">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-212">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-212">1.1</span></span>|
|[<span data-ttu-id="cb2e9-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-215">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="cb2e9-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="cb2e9-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-218">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-218">Type</span></span>

*   [<span data-ttu-id="cb2e9-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="cb2e9-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-220">Requirements</span></span>

|<span data-ttu-id="cb2e9-221">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-221">Requirement</span></span>| <span data-ttu-id="cb2e9-222">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-224">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-224">1.1</span></span>|
|[<span data-ttu-id="cb2e9-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-227">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="cb2e9-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="cb2e9-229">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-230">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-230">Type</span></span>

*   [<span data-ttu-id="cb2e9-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cb2e9-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-232">Requirements</span></span>

|<span data-ttu-id="cb2e9-233">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-233">Requirement</span></span>| <span data-ttu-id="cb2e9-234">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-236">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-236">1.1</span></span>|
|[<span data-ttu-id="cb2e9-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb2e9-239">例</span><span class="sxs-lookup"><span data-stu-id="cb2e9-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="cb2e9-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="cb2e9-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="cb2e9-242">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-243">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-243">Type</span></span>

*   [<span data-ttu-id="cb2e9-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cb2e9-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-245">Requirements</span></span>

|<span data-ttu-id="cb2e9-246">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-246">Requirement</span></span>| <span data-ttu-id="cb2e9-247">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-249">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-249">1.1</span></span>|
|[<span data-ttu-id="cb2e9-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="cb2e9-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="cb2e9-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="cb2e9-251">Restricted</span></span>|
|[<span data-ttu-id="cb2e9-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="cb2e9-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="cb2e9-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="cb2e9-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="cb2e9-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="cb2e9-256">型</span><span class="sxs-lookup"><span data-stu-id="cb2e9-256">Type</span></span>

*   [<span data-ttu-id="cb2e9-257">UI</span><span class="sxs-lookup"><span data-stu-id="cb2e9-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="cb2e9-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="cb2e9-258">Requirements</span></span>

|<span data-ttu-id="cb2e9-259">要件</span><span class="sxs-lookup"><span data-stu-id="cb2e9-259">Requirement</span></span>| <span data-ttu-id="cb2e9-260">値</span><span class="sxs-lookup"><span data-stu-id="cb2e9-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb2e9-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="cb2e9-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cb2e9-262">1.1</span><span class="sxs-lookup"><span data-stu-id="cb2e9-262">1.1</span></span>|
|[<span data-ttu-id="cb2e9-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="cb2e9-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cb2e9-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="cb2e9-264">Compose or Read</span></span>|
