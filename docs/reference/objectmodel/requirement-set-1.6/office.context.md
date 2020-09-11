---
title: Office コンテキスト要件セット1.6
description: メールボックス API 要件セット1.6 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e8cfb6992b8a654a8f348a61ad8d581ffe887df5
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430577"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="ea3f6-103">コンテキスト (メールボックス要件セット 1.6)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ea3f6-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ea3f6-104">[Office](office.md).context</span></span>

<span data-ttu-id="ea3f6-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ea3f6-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea3f6-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-107">Requirements</span></span>

|<span data-ttu-id="ea3f6-108">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-108">Requirement</span></span>| <span data-ttu-id="ea3f6-109">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-111">1.1</span></span>|
|[<span data-ttu-id="ea3f6-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ea3f6-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ea3f6-114">Properties</span></span>

| <span data-ttu-id="ea3f6-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="ea3f6-115">Property</span></span> | <span data-ttu-id="ea3f6-116">モード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-116">Modes</span></span> | <span data-ttu-id="ea3f6-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-117">Return type</span></span> | <span data-ttu-id="ea3f6-118">最小値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-118">Minimum</span></span><br><span data-ttu-id="ea3f6-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="ea3f6-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ea3f6-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ea3f6-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ea3f6-121">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-121">Compose</span></span><br><span data-ttu-id="ea3f6-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-122">Read</span></span> | <span data-ttu-id="ea3f6-123">文字列</span><span class="sxs-lookup"><span data-stu-id="ea3f6-123">String</span></span> | [<span data-ttu-id="ea3f6-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-125">ダン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ea3f6-126">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-126">Compose</span></span><br><span data-ttu-id="ea3f6-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-127">Read</span></span> | [<span data-ttu-id="ea3f6-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ea3f6-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ea3f6-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ea3f6-131">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-131">Compose</span></span><br><span data-ttu-id="ea3f6-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-132">Read</span></span> | <span data-ttu-id="ea3f6-133">文字列</span><span class="sxs-lookup"><span data-stu-id="ea3f6-133">String</span></span> | [<span data-ttu-id="ea3f6-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-135">主催</span><span class="sxs-lookup"><span data-stu-id="ea3f6-135">host</span></span>](#host-hosttype) | <span data-ttu-id="ea3f6-136">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-136">Compose</span></span><br><span data-ttu-id="ea3f6-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-137">Read</span></span> | [<span data-ttu-id="ea3f6-138">HostType</span><span class="sxs-lookup"><span data-stu-id="ea3f6-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="ea3f6-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ea3f6-141">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-141">Compose</span></span><br><span data-ttu-id="ea3f6-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-142">Read</span></span> | [<span data-ttu-id="ea3f6-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="ea3f6-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ea3f6-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ea3f6-146">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-146">Compose</span></span><br><span data-ttu-id="ea3f6-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-147">Read</span></span> | [<span data-ttu-id="ea3f6-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ea3f6-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-150">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ea3f6-151">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-151">Compose</span></span><br><span data-ttu-id="ea3f6-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-152">Read</span></span> | [<span data-ttu-id="ea3f6-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ea3f6-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ea3f6-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ea3f6-156">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-156">Compose</span></span><br><span data-ttu-id="ea3f6-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-157">Read</span></span> | [<span data-ttu-id="ea3f6-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ea3f6-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ea3f6-160">UI</span><span class="sxs-lookup"><span data-stu-id="ea3f6-160">ui</span></span>](#ui-ui) | <span data-ttu-id="ea3f6-161">作成</span><span class="sxs-lookup"><span data-stu-id="ea3f6-161">Compose</span></span><br><span data-ttu-id="ea3f6-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="ea3f6-162">Read</span></span> | [<span data-ttu-id="ea3f6-163">UI</span><span class="sxs-lookup"><span data-stu-id="ea3f6-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="ea3f6-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ea3f6-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="ea3f6-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ea3f6-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ea3f6-166">contentLanguage: String</span></span>

<span data-ttu-id="ea3f6-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ea3f6-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-169">型</span><span class="sxs-lookup"><span data-stu-id="ea3f6-169">Type</span></span>

*   <span data-ttu-id="ea3f6-170">String</span><span class="sxs-lookup"><span data-stu-id="ea3f6-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea3f6-171">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-171">Requirements</span></span>

|<span data-ttu-id="ea3f6-172">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-172">Requirement</span></span>| <span data-ttu-id="ea3f6-173">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-175">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-175">1.1</span></span>|
|[<span data-ttu-id="ea3f6-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-178">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ea3f6-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ea3f6-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-181">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-181">Type</span></span>

*   [<span data-ttu-id="ea3f6-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ea3f6-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-183">Requirements</span></span>

|<span data-ttu-id="ea3f6-184">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-184">Requirement</span></span>| <span data-ttu-id="ea3f6-185">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-187">1.1</span></span>|
|[<span data-ttu-id="ea3f6-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-190">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ea3f6-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ea3f6-191">displayLanguage: String</span></span>

<span data-ttu-id="ea3f6-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ea3f6-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-194">型</span><span class="sxs-lookup"><span data-stu-id="ea3f6-194">Type</span></span>

*   <span data-ttu-id="ea3f6-195">String</span><span class="sxs-lookup"><span data-stu-id="ea3f6-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ea3f6-196">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-196">Requirements</span></span>

|<span data-ttu-id="ea3f6-197">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-197">Requirement</span></span>| <span data-ttu-id="ea3f6-198">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-200">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-200">1.1</span></span>|
|[<span data-ttu-id="ea3f6-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-203">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ea3f6-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ea3f6-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-206">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-206">Type</span></span>

*   [<span data-ttu-id="ea3f6-207">HostType</span><span class="sxs-lookup"><span data-stu-id="ea3f6-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-208">Requirements</span></span>

|<span data-ttu-id="ea3f6-209">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-209">Requirement</span></span>| <span data-ttu-id="ea3f6-210">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-212">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-212">1.1</span></span>|
|[<span data-ttu-id="ea3f6-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-215">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="ea3f6-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ea3f6-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-218">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-218">Type</span></span>

*   [<span data-ttu-id="ea3f6-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ea3f6-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-220">Requirements</span></span>

|<span data-ttu-id="ea3f6-221">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-221">Requirement</span></span>| <span data-ttu-id="ea3f6-222">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-224">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-224">1.1</span></span>|
|[<span data-ttu-id="ea3f6-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-227">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ea3f6-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ea3f6-229">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-230">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-230">Type</span></span>

*   [<span data-ttu-id="ea3f6-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ea3f6-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-232">Requirements</span></span>

|<span data-ttu-id="ea3f6-233">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-233">Requirement</span></span>| <span data-ttu-id="ea3f6-234">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-236">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-236">1.1</span></span>|
|[<span data-ttu-id="ea3f6-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ea3f6-239">例</span><span class="sxs-lookup"><span data-stu-id="ea3f6-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ea3f6-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ea3f6-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ea3f6-242">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-243">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-243">Type</span></span>

*   [<span data-ttu-id="ea3f6-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ea3f6-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-245">Requirements</span></span>

|<span data-ttu-id="ea3f6-246">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-246">Requirement</span></span>| <span data-ttu-id="ea3f6-247">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-249">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-249">1.1</span></span>|
|[<span data-ttu-id="ea3f6-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="ea3f6-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ea3f6-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="ea3f6-251">Restricted</span></span>|
|[<span data-ttu-id="ea3f6-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ea3f6-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ea3f6-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ea3f6-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="ea3f6-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ea3f6-256">種類</span><span class="sxs-lookup"><span data-stu-id="ea3f6-256">Type</span></span>

*   [<span data-ttu-id="ea3f6-257">UI</span><span class="sxs-lookup"><span data-stu-id="ea3f6-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ea3f6-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="ea3f6-258">Requirements</span></span>

|<span data-ttu-id="ea3f6-259">要件</span><span class="sxs-lookup"><span data-stu-id="ea3f6-259">Requirement</span></span>| <span data-ttu-id="ea3f6-260">値</span><span class="sxs-lookup"><span data-stu-id="ea3f6-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="ea3f6-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="ea3f6-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ea3f6-262">1.1</span><span class="sxs-lookup"><span data-stu-id="ea3f6-262">1.1</span></span>|
|[<span data-ttu-id="ea3f6-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="ea3f6-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ea3f6-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="ea3f6-264">Compose or Read</span></span>|
