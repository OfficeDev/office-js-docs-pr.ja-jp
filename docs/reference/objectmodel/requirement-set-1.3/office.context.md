---
title: Office コンテキスト要件セット1.3
description: メールボックス API 要件セット1.3 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: abfd445e65660daacb3b0058466332c4a08c3889
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293745"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="2b1bc-103">コンテキスト (メールボックス要件セット 1.3)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="2b1bc-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="2b1bc-104">[Office](office.md).context</span></span>

<span data-ttu-id="2b1bc-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="2b1bc-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.3)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b1bc-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-107">Requirements</span></span>

|<span data-ttu-id="2b1bc-108">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-108">Requirement</span></span>| <span data-ttu-id="2b1bc-109">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-111">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-111">1.1</span></span>|
|[<span data-ttu-id="2b1bc-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2b1bc-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2b1bc-114">Properties</span></span>

| <span data-ttu-id="2b1bc-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="2b1bc-115">Property</span></span> | <span data-ttu-id="2b1bc-116">モード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-116">Modes</span></span> | <span data-ttu-id="2b1bc-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="2b1bc-117">Return type</span></span> | <span data-ttu-id="2b1bc-118">最小値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-118">Minimum</span></span><br><span data-ttu-id="2b1bc-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="2b1bc-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2b1bc-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="2b1bc-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="2b1bc-121">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-121">Compose</span></span><br><span data-ttu-id="2b1bc-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-122">Read</span></span> | <span data-ttu-id="2b1bc-123">String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-123">String</span></span> | [<span data-ttu-id="2b1bc-124">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-125">ダン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="2b1bc-126">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-126">Compose</span></span><br><span data-ttu-id="2b1bc-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-127">Read</span></span> | [<span data-ttu-id="2b1bc-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2b1bc-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-129">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="2b1bc-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="2b1bc-131">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-131">Compose</span></span><br><span data-ttu-id="2b1bc-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-132">Read</span></span> | <span data-ttu-id="2b1bc-133">String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-133">String</span></span> | [<span data-ttu-id="2b1bc-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-135">主催</span><span class="sxs-lookup"><span data-stu-id="2b1bc-135">host</span></span>](#host-hosttype) | <span data-ttu-id="2b1bc-136">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-136">Compose</span></span><br><span data-ttu-id="2b1bc-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-137">Read</span></span> | [<span data-ttu-id="2b1bc-138">HostType</span><span class="sxs-lookup"><span data-stu-id="2b1bc-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="2b1bc-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="2b1bc-141">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-141">Compose</span></span><br><span data-ttu-id="2b1bc-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-142">Read</span></span> | [<span data-ttu-id="2b1bc-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="2b1bc-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-144">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-145">platform</span><span class="sxs-lookup"><span data-stu-id="2b1bc-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="2b1bc-146">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-146">Compose</span></span><br><span data-ttu-id="2b1bc-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-147">Read</span></span> | [<span data-ttu-id="2b1bc-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2b1bc-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-150">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="2b1bc-151">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-151">Compose</span></span><br><span data-ttu-id="2b1bc-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-152">Read</span></span> | [<span data-ttu-id="2b1bc-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2b1bc-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-154">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="2b1bc-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="2b1bc-156">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-156">Compose</span></span><br><span data-ttu-id="2b1bc-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-157">Read</span></span> | [<span data-ttu-id="2b1bc-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2b1bc-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-159">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2b1bc-160">UI</span><span class="sxs-lookup"><span data-stu-id="2b1bc-160">ui</span></span>](#ui-ui) | <span data-ttu-id="2b1bc-161">作成</span><span class="sxs-lookup"><span data-stu-id="2b1bc-161">Compose</span></span><br><span data-ttu-id="2b1bc-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="2b1bc-162">Read</span></span> | [<span data-ttu-id="2b1bc-163">UI</span><span class="sxs-lookup"><span data-stu-id="2b1bc-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="2b1bc-164">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="2b1bc-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="2b1bc-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="2b1bc-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-166">contentLanguage: String</span></span>

<span data-ttu-id="2b1bc-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="2b1bc-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-169">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-169">Type</span></span>

*   <span data-ttu-id="2b1bc-170">String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b1bc-171">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-171">Requirements</span></span>

|<span data-ttu-id="2b1bc-172">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-172">Requirement</span></span>| <span data-ttu-id="2b1bc-173">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-175">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-175">1.1</span></span>|
|[<span data-ttu-id="2b1bc-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-178">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="2b1bc-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="2b1bc-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-181">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-181">Type</span></span>

*   [<span data-ttu-id="2b1bc-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2b1bc-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-183">Requirements</span></span>

|<span data-ttu-id="2b1bc-184">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-184">Requirement</span></span>| <span data-ttu-id="2b1bc-185">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-187">1.1</span></span>|
|[<span data-ttu-id="2b1bc-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-190">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="2b1bc-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-191">displayLanguage: String</span></span>

<span data-ttu-id="2b1bc-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="2b1bc-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-194">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-194">Type</span></span>

*   <span data-ttu-id="2b1bc-195">String</span><span class="sxs-lookup"><span data-stu-id="2b1bc-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b1bc-196">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-196">Requirements</span></span>

|<span data-ttu-id="2b1bc-197">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-197">Requirement</span></span>| <span data-ttu-id="2b1bc-198">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-200">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-200">1.1</span></span>|
|[<span data-ttu-id="2b1bc-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-203">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="2b1bc-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="2b1bc-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-206">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-206">Type</span></span>

*   [<span data-ttu-id="2b1bc-207">HostType</span><span class="sxs-lookup"><span data-stu-id="2b1bc-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-208">Requirements</span></span>

|<span data-ttu-id="2b1bc-209">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-209">Requirement</span></span>| <span data-ttu-id="2b1bc-210">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-212">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-212">1.1</span></span>|
|[<span data-ttu-id="2b1bc-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-215">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="2b1bc-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="2b1bc-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-218">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-218">Type</span></span>

*   [<span data-ttu-id="2b1bc-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2b1bc-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-220">Requirements</span></span>

|<span data-ttu-id="2b1bc-221">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-221">Requirement</span></span>| <span data-ttu-id="2b1bc-222">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-224">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-224">1.1</span></span>|
|[<span data-ttu-id="2b1bc-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-227">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="2b1bc-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="2b1bc-229">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-230">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-230">Type</span></span>

*   [<span data-ttu-id="2b1bc-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2b1bc-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-232">Requirements</span></span>

|<span data-ttu-id="2b1bc-233">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-233">Requirement</span></span>| <span data-ttu-id="2b1bc-234">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-236">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-236">1.1</span></span>|
|[<span data-ttu-id="2b1bc-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b1bc-239">例</span><span class="sxs-lookup"><span data-stu-id="2b1bc-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="2b1bc-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="2b1bc-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="2b1bc-242">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-243">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-243">Type</span></span>

*   [<span data-ttu-id="2b1bc-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2b1bc-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-245">Requirements</span></span>

|<span data-ttu-id="2b1bc-246">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-246">Requirement</span></span>| <span data-ttu-id="2b1bc-247">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-249">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-249">1.1</span></span>|
|[<span data-ttu-id="2b1bc-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="2b1bc-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="2b1bc-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="2b1bc-251">Restricted</span></span>|
|[<span data-ttu-id="2b1bc-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="2b1bc-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="2b1bc-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="2b1bc-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="2b1bc-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="2b1bc-256">型</span><span class="sxs-lookup"><span data-stu-id="2b1bc-256">Type</span></span>

*   [<span data-ttu-id="2b1bc-257">UI</span><span class="sxs-lookup"><span data-stu-id="2b1bc-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="2b1bc-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="2b1bc-258">Requirements</span></span>

|<span data-ttu-id="2b1bc-259">要件</span><span class="sxs-lookup"><span data-stu-id="2b1bc-259">Requirement</span></span>| <span data-ttu-id="2b1bc-260">値</span><span class="sxs-lookup"><span data-stu-id="2b1bc-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b1bc-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="2b1bc-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2b1bc-262">1.1</span><span class="sxs-lookup"><span data-stu-id="2b1bc-262">1.1</span></span>|
|[<span data-ttu-id="2b1bc-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="2b1bc-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2b1bc-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="2b1bc-264">Compose or Read</span></span>|
