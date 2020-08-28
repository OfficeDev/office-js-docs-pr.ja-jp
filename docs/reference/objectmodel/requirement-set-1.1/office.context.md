---
title: Office コンテキスト要件セット1.1
description: メールボックス API 要件セット1.1 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f1788ecfb832b70fd3622df3f17a2549f897fa84
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293787"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="3dfd7-103">コンテキスト (メールボックス要件セット 1.1)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="3dfd7-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="3dfd7-104">[Office](office.md).context</span></span>

<span data-ttu-id="3dfd7-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="3dfd7-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dfd7-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-107">Requirements</span></span>

|<span data-ttu-id="3dfd7-108">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-108">Requirement</span></span>| <span data-ttu-id="3dfd7-109">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-111">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-111">1.1</span></span>|
|[<span data-ttu-id="3dfd7-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3dfd7-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3dfd7-114">Properties</span></span>

| <span data-ttu-id="3dfd7-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3dfd7-115">Property</span></span> | <span data-ttu-id="3dfd7-116">モード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-116">Modes</span></span> | <span data-ttu-id="3dfd7-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3dfd7-117">Return type</span></span> | <span data-ttu-id="3dfd7-118">最小値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-118">Minimum</span></span><br><span data-ttu-id="3dfd7-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="3dfd7-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3dfd7-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="3dfd7-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="3dfd7-121">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-121">Compose</span></span><br><span data-ttu-id="3dfd7-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-122">Read</span></span> | <span data-ttu-id="3dfd7-123">String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-123">String</span></span> | [<span data-ttu-id="3dfd7-124">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-125">ダン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="3dfd7-126">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-126">Compose</span></span><br><span data-ttu-id="3dfd7-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-127">Read</span></span> | [<span data-ttu-id="3dfd7-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3dfd7-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-129">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3dfd7-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3dfd7-131">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-131">Compose</span></span><br><span data-ttu-id="3dfd7-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-132">Read</span></span> | <span data-ttu-id="3dfd7-133">String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-133">String</span></span> | [<span data-ttu-id="3dfd7-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-135">主催</span><span class="sxs-lookup"><span data-stu-id="3dfd7-135">host</span></span>](#host-hosttype) | <span data-ttu-id="3dfd7-136">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-136">Compose</span></span><br><span data-ttu-id="3dfd7-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-137">Read</span></span> | [<span data-ttu-id="3dfd7-138">HostType</span><span class="sxs-lookup"><span data-stu-id="3dfd7-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="3dfd7-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="3dfd7-141">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-141">Compose</span></span><br><span data-ttu-id="3dfd7-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-142">Read</span></span> | [<span data-ttu-id="3dfd7-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="3dfd7-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-144">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-145">platform</span><span class="sxs-lookup"><span data-stu-id="3dfd7-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="3dfd7-146">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-146">Compose</span></span><br><span data-ttu-id="3dfd7-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-147">Read</span></span> | [<span data-ttu-id="3dfd7-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3dfd7-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-150">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="3dfd7-151">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-151">Compose</span></span><br><span data-ttu-id="3dfd7-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-152">Read</span></span> | [<span data-ttu-id="3dfd7-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3dfd7-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-154">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dfd7-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="3dfd7-156">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-156">Compose</span></span><br><span data-ttu-id="3dfd7-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-157">Read</span></span> | [<span data-ttu-id="3dfd7-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dfd7-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-159">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3dfd7-160">UI</span><span class="sxs-lookup"><span data-stu-id="3dfd7-160">ui</span></span>](#ui-ui) | <span data-ttu-id="3dfd7-161">作成</span><span class="sxs-lookup"><span data-stu-id="3dfd7-161">Compose</span></span><br><span data-ttu-id="3dfd7-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="3dfd7-162">Read</span></span> | [<span data-ttu-id="3dfd7-163">UI</span><span class="sxs-lookup"><span data-stu-id="3dfd7-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1) | [<span data-ttu-id="3dfd7-164">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="3dfd7-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="3dfd7-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="3dfd7-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-166">contentLanguage: String</span></span>

<span data-ttu-id="3dfd7-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="3dfd7-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-169">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-169">Type</span></span>

*   <span data-ttu-id="3dfd7-170">String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dfd7-171">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-171">Requirements</span></span>

|<span data-ttu-id="3dfd7-172">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-172">Requirement</span></span>| <span data-ttu-id="3dfd7-173">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-175">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-175">1.1</span></span>|
|[<span data-ttu-id="3dfd7-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-178">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="3dfd7-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="3dfd7-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-181">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-181">Type</span></span>

*   [<span data-ttu-id="3dfd7-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3dfd7-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-183">Requirements</span></span>

|<span data-ttu-id="3dfd7-184">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-184">Requirement</span></span>| <span data-ttu-id="3dfd7-185">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-187">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-187">1.1</span></span>|
|[<span data-ttu-id="3dfd7-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-190">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="3dfd7-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-191">displayLanguage: String</span></span>

<span data-ttu-id="3dfd7-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="3dfd7-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-194">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-194">Type</span></span>

*   <span data-ttu-id="3dfd7-195">String</span><span class="sxs-lookup"><span data-stu-id="3dfd7-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3dfd7-196">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-196">Requirements</span></span>

|<span data-ttu-id="3dfd7-197">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-197">Requirement</span></span>| <span data-ttu-id="3dfd7-198">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-200">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-200">1.1</span></span>|
|[<span data-ttu-id="3dfd7-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-203">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="3dfd7-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="3dfd7-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-206">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-206">Type</span></span>

*   [<span data-ttu-id="3dfd7-207">HostType</span><span class="sxs-lookup"><span data-stu-id="3dfd7-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-208">Requirements</span></span>

|<span data-ttu-id="3dfd7-209">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-209">Requirement</span></span>| <span data-ttu-id="3dfd7-210">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-212">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-212">1.1</span></span>|
|[<span data-ttu-id="3dfd7-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-215">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="3dfd7-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="3dfd7-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-218">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-218">Type</span></span>

*   [<span data-ttu-id="3dfd7-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3dfd7-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-220">Requirements</span></span>

|<span data-ttu-id="3dfd7-221">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-221">Requirement</span></span>| <span data-ttu-id="3dfd7-222">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-224">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-224">1.1</span></span>|
|[<span data-ttu-id="3dfd7-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-227">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="3dfd7-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="3dfd7-229">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-230">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-230">Type</span></span>

*   [<span data-ttu-id="3dfd7-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3dfd7-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-232">Requirements</span></span>

|<span data-ttu-id="3dfd7-233">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-233">Requirement</span></span>| <span data-ttu-id="3dfd7-234">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-236">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-236">1.1</span></span>|
|[<span data-ttu-id="3dfd7-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3dfd7-239">例</span><span class="sxs-lookup"><span data-stu-id="3dfd7-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="3dfd7-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="3dfd7-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3dfd7-242">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-243">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-243">Type</span></span>

*   [<span data-ttu-id="3dfd7-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3dfd7-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-245">Requirements</span></span>

|<span data-ttu-id="3dfd7-246">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-246">Requirement</span></span>| <span data-ttu-id="3dfd7-247">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-249">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-249">1.1</span></span>|
|[<span data-ttu-id="3dfd7-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3dfd7-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="3dfd7-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="3dfd7-251">Restricted</span></span>|
|[<span data-ttu-id="3dfd7-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="3dfd7-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="3dfd7-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="3dfd7-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="3dfd7-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3dfd7-256">型</span><span class="sxs-lookup"><span data-stu-id="3dfd7-256">Type</span></span>

*   [<span data-ttu-id="3dfd7-257">UI</span><span class="sxs-lookup"><span data-stu-id="3dfd7-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="3dfd7-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="3dfd7-258">Requirements</span></span>

|<span data-ttu-id="3dfd7-259">要件</span><span class="sxs-lookup"><span data-stu-id="3dfd7-259">Requirement</span></span>| <span data-ttu-id="3dfd7-260">値</span><span class="sxs-lookup"><span data-stu-id="3dfd7-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="3dfd7-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3dfd7-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3dfd7-262">1.1</span><span class="sxs-lookup"><span data-stu-id="3dfd7-262">1.1</span></span>|
|[<span data-ttu-id="3dfd7-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3dfd7-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3dfd7-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3dfd7-264">Compose or Read</span></span>|
