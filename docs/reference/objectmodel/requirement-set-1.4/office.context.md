---
title: Office コンテキスト要件セット1.4
description: メールボックス API 要件セット1.4 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: cda0fc55fa4224f8bd5f30c80e43febad5478eb3
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430731"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="80ce3-103">コンテキスト (メールボックス要件セット 1.4)</span><span class="sxs-lookup"><span data-stu-id="80ce3-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="80ce3-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="80ce3-104">[Office](office.md).context</span></span>

<span data-ttu-id="80ce3-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="80ce3-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="80ce3-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="80ce3-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-107">Requirements</span></span>

|<span data-ttu-id="80ce3-108">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-108">Requirement</span></span>| <span data-ttu-id="80ce3-109">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-111">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-111">1.1</span></span>|
|[<span data-ttu-id="80ce3-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="80ce3-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="80ce3-114">Properties</span></span>

| <span data-ttu-id="80ce3-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="80ce3-115">Property</span></span> | <span data-ttu-id="80ce3-116">モード</span><span class="sxs-lookup"><span data-stu-id="80ce3-116">Modes</span></span> | <span data-ttu-id="80ce3-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-117">Return type</span></span> | <span data-ttu-id="80ce3-118">最小値</span><span class="sxs-lookup"><span data-stu-id="80ce3-118">Minimum</span></span><br><span data-ttu-id="80ce3-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="80ce3-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="80ce3-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="80ce3-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="80ce3-121">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-121">Compose</span></span><br><span data-ttu-id="80ce3-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-122">Read</span></span> | <span data-ttu-id="80ce3-123">文字列</span><span class="sxs-lookup"><span data-stu-id="80ce3-123">String</span></span> | [<span data-ttu-id="80ce3-124">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-125">ダン</span><span class="sxs-lookup"><span data-stu-id="80ce3-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="80ce3-126">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-126">Compose</span></span><br><span data-ttu-id="80ce3-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-127">Read</span></span> | [<span data-ttu-id="80ce3-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="80ce3-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-129">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="80ce3-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="80ce3-131">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-131">Compose</span></span><br><span data-ttu-id="80ce3-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-132">Read</span></span> | <span data-ttu-id="80ce3-133">文字列</span><span class="sxs-lookup"><span data-stu-id="80ce3-133">String</span></span> | [<span data-ttu-id="80ce3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-135">主催</span><span class="sxs-lookup"><span data-stu-id="80ce3-135">host</span></span>](#host-hosttype) | <span data-ttu-id="80ce3-136">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-136">Compose</span></span><br><span data-ttu-id="80ce3-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-137">Read</span></span> | [<span data-ttu-id="80ce3-138">HostType</span><span class="sxs-lookup"><span data-stu-id="80ce3-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="80ce3-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="80ce3-141">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-141">Compose</span></span><br><span data-ttu-id="80ce3-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-142">Read</span></span> | [<span data-ttu-id="80ce3-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="80ce3-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-144">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="80ce3-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="80ce3-146">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-146">Compose</span></span><br><span data-ttu-id="80ce3-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-147">Read</span></span> | [<span data-ttu-id="80ce3-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="80ce3-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-150">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="80ce3-151">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-151">Compose</span></span><br><span data-ttu-id="80ce3-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-152">Read</span></span> | [<span data-ttu-id="80ce3-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="80ce3-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-154">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="80ce3-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="80ce3-156">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-156">Compose</span></span><br><span data-ttu-id="80ce3-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-157">Read</span></span> | [<span data-ttu-id="80ce3-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="80ce3-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-159">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="80ce3-160">UI</span><span class="sxs-lookup"><span data-stu-id="80ce3-160">ui</span></span>](#ui-ui) | <span data-ttu-id="80ce3-161">作成</span><span class="sxs-lookup"><span data-stu-id="80ce3-161">Compose</span></span><br><span data-ttu-id="80ce3-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="80ce3-162">Read</span></span> | [<span data-ttu-id="80ce3-163">UI</span><span class="sxs-lookup"><span data-stu-id="80ce3-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="80ce3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="80ce3-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="80ce3-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="80ce3-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="80ce3-166">contentLanguage: String</span></span>

<span data-ttu-id="80ce3-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="80ce3-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-169">型</span><span class="sxs-lookup"><span data-stu-id="80ce3-169">Type</span></span>

*   <span data-ttu-id="80ce3-170">String</span><span class="sxs-lookup"><span data-stu-id="80ce3-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80ce3-171">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-171">Requirements</span></span>

|<span data-ttu-id="80ce3-172">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-172">Requirement</span></span>| <span data-ttu-id="80ce3-173">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-175">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-175">1.1</span></span>|
|[<span data-ttu-id="80ce3-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-178">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="80ce3-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="80ce3-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="80ce3-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-181">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-181">Type</span></span>

*   [<span data-ttu-id="80ce3-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="80ce3-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="80ce3-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-183">Requirements</span></span>

|<span data-ttu-id="80ce3-184">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-184">Requirement</span></span>| <span data-ttu-id="80ce3-185">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-187">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-187">1.1</span></span>|
|[<span data-ttu-id="80ce3-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-190">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="80ce3-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="80ce3-191">displayLanguage: String</span></span>

<span data-ttu-id="80ce3-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="80ce3-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**表示言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-194">型</span><span class="sxs-lookup"><span data-stu-id="80ce3-194">Type</span></span>

*   <span data-ttu-id="80ce3-195">String</span><span class="sxs-lookup"><span data-stu-id="80ce3-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80ce3-196">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-196">Requirements</span></span>

|<span data-ttu-id="80ce3-197">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-197">Requirement</span></span>| <span data-ttu-id="80ce3-198">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-200">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-200">1.1</span></span>|
|[<span data-ttu-id="80ce3-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-203">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="80ce3-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="80ce3-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="80ce3-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-206">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-206">Type</span></span>

*   [<span data-ttu-id="80ce3-207">HostType</span><span class="sxs-lookup"><span data-stu-id="80ce3-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="80ce3-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-208">Requirements</span></span>

|<span data-ttu-id="80ce3-209">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-209">Requirement</span></span>| <span data-ttu-id="80ce3-210">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-212">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-212">1.1</span></span>|
|[<span data-ttu-id="80ce3-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-215">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="80ce3-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="80ce3-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="80ce3-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-218">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-218">Type</span></span>

*   [<span data-ttu-id="80ce3-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="80ce3-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="80ce3-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-220">Requirements</span></span>

|<span data-ttu-id="80ce3-221">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-221">Requirement</span></span>| <span data-ttu-id="80ce3-222">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-224">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-224">1.1</span></span>|
|[<span data-ttu-id="80ce3-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-227">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="80ce3-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="80ce3-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="80ce3-229">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-230">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-230">Type</span></span>

*   [<span data-ttu-id="80ce3-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="80ce3-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="80ce3-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-232">Requirements</span></span>

|<span data-ttu-id="80ce3-233">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-233">Requirement</span></span>| <span data-ttu-id="80ce3-234">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-236">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-236">1.1</span></span>|
|[<span data-ttu-id="80ce3-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80ce3-239">例</span><span class="sxs-lookup"><span data-stu-id="80ce3-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="80ce3-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="80ce3-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="80ce3-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="80ce3-242">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="80ce3-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-243">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-243">Type</span></span>

*   [<span data-ttu-id="80ce3-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="80ce3-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="80ce3-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-245">Requirements</span></span>

|<span data-ttu-id="80ce3-246">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-246">Requirement</span></span>| <span data-ttu-id="80ce3-247">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-249">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-249">1.1</span></span>|
|[<span data-ttu-id="80ce3-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="80ce3-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="80ce3-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="80ce3-251">Restricted</span></span>|
|[<span data-ttu-id="80ce3-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="80ce3-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="80ce3-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="80ce3-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="80ce3-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="80ce3-256">種類</span><span class="sxs-lookup"><span data-stu-id="80ce3-256">Type</span></span>

*   [<span data-ttu-id="80ce3-257">UI</span><span class="sxs-lookup"><span data-stu-id="80ce3-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="80ce3-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="80ce3-258">Requirements</span></span>

|<span data-ttu-id="80ce3-259">要件</span><span class="sxs-lookup"><span data-stu-id="80ce3-259">Requirement</span></span>| <span data-ttu-id="80ce3-260">値</span><span class="sxs-lookup"><span data-stu-id="80ce3-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="80ce3-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="80ce3-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="80ce3-262">1.1</span><span class="sxs-lookup"><span data-stu-id="80ce3-262">1.1</span></span>|
|[<span data-ttu-id="80ce3-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="80ce3-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="80ce3-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="80ce3-264">Compose or Read</span></span>|
