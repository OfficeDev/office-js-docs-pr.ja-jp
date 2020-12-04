---
title: Office コンテキスト要件セット1.7
description: メールボックス API 要件セット1.7 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 4a1ca6b4975ffba2c2bd400267fbe7db63f88244
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570731"
---
# <a name="context-mailbox-requirement-set-17"></a><span data-ttu-id="6cb2a-103">コンテキスト (メールボックス要件セット 1.7)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-103">context (Mailbox requirement set 1.7)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6cb2a-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6cb2a-104">[Office](office.md).context</span></span>

<span data-ttu-id="6cb2a-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6cb2a-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb2a-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-107">Requirements</span></span>

|<span data-ttu-id="6cb2a-108">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-108">Requirement</span></span>| <span data-ttu-id="6cb2a-109">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-111">1.1</span></span>|
|[<span data-ttu-id="6cb2a-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6cb2a-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6cb2a-114">Properties</span></span>

| <span data-ttu-id="6cb2a-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="6cb2a-115">Property</span></span> | <span data-ttu-id="6cb2a-116">モード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-116">Modes</span></span> | <span data-ttu-id="6cb2a-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-117">Return type</span></span> | <span data-ttu-id="6cb2a-118">最小値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-118">Minimum</span></span><br><span data-ttu-id="6cb2a-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="6cb2a-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6cb2a-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6cb2a-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6cb2a-121">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-121">Compose</span></span><br><span data-ttu-id="6cb2a-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-122">Read</span></span> | <span data-ttu-id="6cb2a-123">文字列</span><span class="sxs-lookup"><span data-stu-id="6cb2a-123">String</span></span> | [<span data-ttu-id="6cb2a-124">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-125">ダン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6cb2a-126">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-126">Compose</span></span><br><span data-ttu-id="6cb2a-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-127">Read</span></span> | [<span data-ttu-id="6cb2a-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6cb2a-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6cb2a-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6cb2a-131">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-131">Compose</span></span><br><span data-ttu-id="6cb2a-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-132">Read</span></span> | <span data-ttu-id="6cb2a-133">文字列</span><span class="sxs-lookup"><span data-stu-id="6cb2a-133">String</span></span> | [<span data-ttu-id="6cb2a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-135">主催</span><span class="sxs-lookup"><span data-stu-id="6cb2a-135">host</span></span>](#host-hosttype) | <span data-ttu-id="6cb2a-136">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-136">Compose</span></span><br><span data-ttu-id="6cb2a-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-137">Read</span></span> | [<span data-ttu-id="6cb2a-138">HostType</span><span class="sxs-lookup"><span data-stu-id="6cb2a-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-139">1.5</span><span class="sxs-lookup"><span data-stu-id="6cb2a-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6cb2a-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="6cb2a-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6cb2a-141">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-141">Compose</span></span><br><span data-ttu-id="6cb2a-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-142">Read</span></span> | [<span data-ttu-id="6cb2a-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="6cb2a-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-145">platform</span><span class="sxs-lookup"><span data-stu-id="6cb2a-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="6cb2a-146">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-146">Compose</span></span><br><span data-ttu-id="6cb2a-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-147">Read</span></span> | [<span data-ttu-id="6cb2a-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6cb2a-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-149">1.5</span><span class="sxs-lookup"><span data-stu-id="6cb2a-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6cb2a-150">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6cb2a-151">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-151">Compose</span></span><br><span data-ttu-id="6cb2a-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-152">Read</span></span> | [<span data-ttu-id="6cb2a-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6cb2a-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-154">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb2a-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6cb2a-156">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-156">Compose</span></span><br><span data-ttu-id="6cb2a-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-157">Read</span></span> | [<span data-ttu-id="6cb2a-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb2a-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-159">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb2a-160">UI</span><span class="sxs-lookup"><span data-stu-id="6cb2a-160">ui</span></span>](#ui-ui) | <span data-ttu-id="6cb2a-161">作成</span><span class="sxs-lookup"><span data-stu-id="6cb2a-161">Compose</span></span><br><span data-ttu-id="6cb2a-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="6cb2a-162">Read</span></span> | [<span data-ttu-id="6cb2a-163">UI</span><span class="sxs-lookup"><span data-stu-id="6cb2a-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="6cb2a-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6cb2a-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="6cb2a-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="6cb2a-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6cb2a-166">contentLanguage: String</span></span>

<span data-ttu-id="6cb2a-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6cb2a-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-169">型</span><span class="sxs-lookup"><span data-stu-id="6cb2a-169">Type</span></span>

*   <span data-ttu-id="6cb2a-170">String</span><span class="sxs-lookup"><span data-stu-id="6cb2a-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb2a-171">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-171">Requirements</span></span>

|<span data-ttu-id="6cb2a-172">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-172">Requirement</span></span>| <span data-ttu-id="6cb2a-173">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-175">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-175">1.1</span></span>|
|[<span data-ttu-id="6cb2a-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-178">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6cb2a-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6cb2a-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-181">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-181">Type</span></span>

*   [<span data-ttu-id="6cb2a-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6cb2a-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-183">Requirements</span></span>

|<span data-ttu-id="6cb2a-184">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-184">Requirement</span></span>| <span data-ttu-id="6cb2a-185">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-187">1.1</span></span>|
|[<span data-ttu-id="6cb2a-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-190">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6cb2a-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6cb2a-191">displayLanguage: String</span></span>

<span data-ttu-id="6cb2a-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="6cb2a-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-194">型</span><span class="sxs-lookup"><span data-stu-id="6cb2a-194">Type</span></span>

*   <span data-ttu-id="6cb2a-195">String</span><span class="sxs-lookup"><span data-stu-id="6cb2a-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb2a-196">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-196">Requirements</span></span>

|<span data-ttu-id="6cb2a-197">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-197">Requirement</span></span>| <span data-ttu-id="6cb2a-198">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-200">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-200">1.1</span></span>|
|[<span data-ttu-id="6cb2a-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-203">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="6cb2a-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="6cb2a-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="6cb2a-206">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してホストを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-207">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-207">Type</span></span>

*   [<span data-ttu-id="6cb2a-208">HostType</span><span class="sxs-lookup"><span data-stu-id="6cb2a-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-209">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-209">Requirements</span></span>

|<span data-ttu-id="6cb2a-210">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-210">Requirement</span></span>| <span data-ttu-id="6cb2a-211">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-213">1.5</span><span class="sxs-lookup"><span data-stu-id="6cb2a-213">1.5</span></span>|
|[<span data-ttu-id="6cb2a-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-215">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-216">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="6cb2a-217">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="6cb2a-218">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="6cb2a-219">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-220">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-220">Type</span></span>

*   [<span data-ttu-id="6cb2a-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6cb2a-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-222">Requirements</span></span>

|<span data-ttu-id="6cb2a-223">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-223">Requirement</span></span>| <span data-ttu-id="6cb2a-224">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-226">1.5</span><span class="sxs-lookup"><span data-stu-id="6cb2a-226">1.5</span></span>|
|[<span data-ttu-id="6cb2a-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-229">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6cb2a-230">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6cb2a-231">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-232">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-232">Type</span></span>

*   [<span data-ttu-id="6cb2a-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6cb2a-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-234">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-234">Requirements</span></span>

|<span data-ttu-id="6cb2a-235">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-235">Requirement</span></span>| <span data-ttu-id="6cb2a-236">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-238">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-238">1.1</span></span>|
|[<span data-ttu-id="6cb2a-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-240">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb2a-241">例</span><span class="sxs-lookup"><span data-stu-id="6cb2a-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6cb2a-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6cb2a-243">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6cb2a-244">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-245">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-245">Type</span></span>

*   [<span data-ttu-id="6cb2a-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb2a-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-247">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-247">Requirements</span></span>

|<span data-ttu-id="6cb2a-248">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-248">Requirement</span></span>| <span data-ttu-id="6cb2a-249">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-251">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-251">1.1</span></span>|
|[<span data-ttu-id="6cb2a-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="6cb2a-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6cb2a-253">制限あり</span><span class="sxs-lookup"><span data-stu-id="6cb2a-253">Restricted</span></span>|
|[<span data-ttu-id="6cb2a-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6cb2a-256">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6cb2a-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6cb2a-257">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="6cb2a-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb2a-258">種類</span><span class="sxs-lookup"><span data-stu-id="6cb2a-258">Type</span></span>

*   [<span data-ttu-id="6cb2a-259">UI</span><span class="sxs-lookup"><span data-stu-id="6cb2a-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6cb2a-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="6cb2a-260">Requirements</span></span>

|<span data-ttu-id="6cb2a-261">要件</span><span class="sxs-lookup"><span data-stu-id="6cb2a-261">Requirement</span></span>| <span data-ttu-id="6cb2a-262">値</span><span class="sxs-lookup"><span data-stu-id="6cb2a-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb2a-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="6cb2a-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb2a-264">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb2a-264">1.1</span></span>|
|[<span data-ttu-id="6cb2a-265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="6cb2a-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb2a-266">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="6cb2a-266">Compose or Read</span></span>|
