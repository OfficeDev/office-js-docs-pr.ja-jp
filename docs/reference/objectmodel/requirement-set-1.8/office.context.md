---
title: Office コンテキスト要件セット1.8
description: メールボックス API 要件セット1.8 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: cf49abb05bbe2e5e7b1d4d178c7749d6e7183d2a
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570766"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="90d72-103">コンテキスト (メールボックス要件セット 1.8)</span><span class="sxs-lookup"><span data-stu-id="90d72-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="90d72-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="90d72-104">[Office](office.md).context</span></span>

<span data-ttu-id="90d72-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="90d72-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="90d72-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="90d72-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="90d72-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-107">Requirements</span></span>

|<span data-ttu-id="90d72-108">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-108">Requirement</span></span>| <span data-ttu-id="90d72-109">値</span><span class="sxs-lookup"><span data-stu-id="90d72-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-111">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-111">1.1</span></span>|
|[<span data-ttu-id="90d72-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="90d72-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="90d72-114">Properties</span></span>

| <span data-ttu-id="90d72-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="90d72-115">Property</span></span> | <span data-ttu-id="90d72-116">モード</span><span class="sxs-lookup"><span data-stu-id="90d72-116">Modes</span></span> | <span data-ttu-id="90d72-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="90d72-117">Return type</span></span> | <span data-ttu-id="90d72-118">最小値</span><span class="sxs-lookup"><span data-stu-id="90d72-118">Minimum</span></span><br><span data-ttu-id="90d72-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="90d72-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="90d72-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="90d72-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="90d72-121">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-121">Compose</span></span><br><span data-ttu-id="90d72-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-122">Read</span></span> | <span data-ttu-id="90d72-123">文字列</span><span class="sxs-lookup"><span data-stu-id="90d72-123">String</span></span> | [<span data-ttu-id="90d72-124">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-125">ダン</span><span class="sxs-lookup"><span data-stu-id="90d72-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="90d72-126">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-126">Compose</span></span><br><span data-ttu-id="90d72-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-127">Read</span></span> | [<span data-ttu-id="90d72-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="90d72-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-129">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="90d72-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="90d72-131">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-131">Compose</span></span><br><span data-ttu-id="90d72-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-132">Read</span></span> | <span data-ttu-id="90d72-133">文字列</span><span class="sxs-lookup"><span data-stu-id="90d72-133">String</span></span> | [<span data-ttu-id="90d72-134">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-135">主催</span><span class="sxs-lookup"><span data-stu-id="90d72-135">host</span></span>](#host-hosttype) | <span data-ttu-id="90d72-136">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-136">Compose</span></span><br><span data-ttu-id="90d72-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-137">Read</span></span> | [<span data-ttu-id="90d72-138">HostType</span><span class="sxs-lookup"><span data-stu-id="90d72-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-139">1.5</span><span class="sxs-lookup"><span data-stu-id="90d72-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="90d72-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="90d72-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="90d72-141">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-141">Compose</span></span><br><span data-ttu-id="90d72-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-142">Read</span></span> | [<span data-ttu-id="90d72-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="90d72-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-144">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-145">platform</span><span class="sxs-lookup"><span data-stu-id="90d72-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="90d72-146">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-146">Compose</span></span><br><span data-ttu-id="90d72-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-147">Read</span></span> | [<span data-ttu-id="90d72-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="90d72-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-149">1.5</span><span class="sxs-lookup"><span data-stu-id="90d72-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="90d72-150">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="90d72-151">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-151">Compose</span></span><br><span data-ttu-id="90d72-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-152">Read</span></span> | [<span data-ttu-id="90d72-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="90d72-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-154">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="90d72-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="90d72-156">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-156">Compose</span></span><br><span data-ttu-id="90d72-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-157">Read</span></span> | [<span data-ttu-id="90d72-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="90d72-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-159">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="90d72-160">UI</span><span class="sxs-lookup"><span data-stu-id="90d72-160">ui</span></span>](#ui-ui) | <span data-ttu-id="90d72-161">作成</span><span class="sxs-lookup"><span data-stu-id="90d72-161">Compose</span></span><br><span data-ttu-id="90d72-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="90d72-162">Read</span></span> | [<span data-ttu-id="90d72-163">UI</span><span class="sxs-lookup"><span data-stu-id="90d72-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="90d72-164">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="90d72-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="90d72-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="90d72-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="90d72-166">contentLanguage: String</span></span>

<span data-ttu-id="90d72-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="90d72-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="90d72-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="90d72-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-169">型</span><span class="sxs-lookup"><span data-stu-id="90d72-169">Type</span></span>

*   <span data-ttu-id="90d72-170">String</span><span class="sxs-lookup"><span data-stu-id="90d72-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="90d72-171">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-171">Requirements</span></span>

|<span data-ttu-id="90d72-172">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-172">Requirement</span></span>| <span data-ttu-id="90d72-173">値</span><span class="sxs-lookup"><span data-stu-id="90d72-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-175">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-175">1.1</span></span>|
|[<span data-ttu-id="90d72-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-178">例</span><span class="sxs-lookup"><span data-stu-id="90d72-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="90d72-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="90d72-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="90d72-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="90d72-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-181">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-181">Type</span></span>

*   [<span data-ttu-id="90d72-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="90d72-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="90d72-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-183">Requirements</span></span>

|<span data-ttu-id="90d72-184">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-184">Requirement</span></span>| <span data-ttu-id="90d72-185">値</span><span class="sxs-lookup"><span data-stu-id="90d72-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-187">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-187">1.1</span></span>|
|[<span data-ttu-id="90d72-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-190">例</span><span class="sxs-lookup"><span data-stu-id="90d72-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="90d72-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="90d72-191">displayLanguage: String</span></span>

<span data-ttu-id="90d72-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="90d72-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="90d72-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="90d72-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-194">型</span><span class="sxs-lookup"><span data-stu-id="90d72-194">Type</span></span>

*   <span data-ttu-id="90d72-195">String</span><span class="sxs-lookup"><span data-stu-id="90d72-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="90d72-196">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-196">Requirements</span></span>

|<span data-ttu-id="90d72-197">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-197">Requirement</span></span>| <span data-ttu-id="90d72-198">値</span><span class="sxs-lookup"><span data-stu-id="90d72-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-200">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-200">1.1</span></span>|
|[<span data-ttu-id="90d72-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-203">例</span><span class="sxs-lookup"><span data-stu-id="90d72-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="90d72-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="90d72-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="90d72-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="90d72-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="90d72-206">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してホストを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="90d72-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-207">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-207">Type</span></span>

*   [<span data-ttu-id="90d72-208">HostType</span><span class="sxs-lookup"><span data-stu-id="90d72-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="90d72-209">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-209">Requirements</span></span>

|<span data-ttu-id="90d72-210">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-210">Requirement</span></span>| <span data-ttu-id="90d72-211">値</span><span class="sxs-lookup"><span data-stu-id="90d72-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-213">1.5</span><span class="sxs-lookup"><span data-stu-id="90d72-213">1.5</span></span>|
|[<span data-ttu-id="90d72-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-215">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-216">例</span><span class="sxs-lookup"><span data-stu-id="90d72-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="90d72-217">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="90d72-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="90d72-218">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="90d72-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="90d72-219">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="90d72-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-220">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-220">Type</span></span>

*   [<span data-ttu-id="90d72-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="90d72-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="90d72-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-222">Requirements</span></span>

|<span data-ttu-id="90d72-223">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-223">Requirement</span></span>| <span data-ttu-id="90d72-224">値</span><span class="sxs-lookup"><span data-stu-id="90d72-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-226">1.5</span><span class="sxs-lookup"><span data-stu-id="90d72-226">1.5</span></span>|
|[<span data-ttu-id="90d72-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-229">例</span><span class="sxs-lookup"><span data-stu-id="90d72-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="90d72-230">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="90d72-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="90d72-231">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="90d72-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-232">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-232">Type</span></span>

*   [<span data-ttu-id="90d72-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="90d72-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="90d72-234">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-234">Requirements</span></span>

|<span data-ttu-id="90d72-235">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-235">Requirement</span></span>| <span data-ttu-id="90d72-236">値</span><span class="sxs-lookup"><span data-stu-id="90d72-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-238">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-238">1.1</span></span>|
|[<span data-ttu-id="90d72-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-240">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="90d72-241">例</span><span class="sxs-lookup"><span data-stu-id="90d72-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="90d72-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="90d72-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="90d72-243">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="90d72-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="90d72-244">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="90d72-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-245">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-245">Type</span></span>

*   [<span data-ttu-id="90d72-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="90d72-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="90d72-247">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-247">Requirements</span></span>

|<span data-ttu-id="90d72-248">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-248">Requirement</span></span>| <span data-ttu-id="90d72-249">値</span><span class="sxs-lookup"><span data-stu-id="90d72-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-251">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-251">1.1</span></span>|
|[<span data-ttu-id="90d72-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="90d72-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="90d72-253">制限あり</span><span class="sxs-lookup"><span data-stu-id="90d72-253">Restricted</span></span>|
|[<span data-ttu-id="90d72-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="90d72-256">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="90d72-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="90d72-257">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="90d72-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="90d72-258">種類</span><span class="sxs-lookup"><span data-stu-id="90d72-258">Type</span></span>

*   [<span data-ttu-id="90d72-259">UI</span><span class="sxs-lookup"><span data-stu-id="90d72-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="90d72-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="90d72-260">Requirements</span></span>

|<span data-ttu-id="90d72-261">要件</span><span class="sxs-lookup"><span data-stu-id="90d72-261">Requirement</span></span>| <span data-ttu-id="90d72-262">値</span><span class="sxs-lookup"><span data-stu-id="90d72-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="90d72-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="90d72-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="90d72-264">1.1</span><span class="sxs-lookup"><span data-stu-id="90d72-264">1.1</span></span>|
|[<span data-ttu-id="90d72-265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="90d72-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="90d72-266">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="90d72-266">Compose or Read</span></span>|
