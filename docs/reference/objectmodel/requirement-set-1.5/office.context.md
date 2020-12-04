---
title: Office.context - 要件セット 1.5
description: メールボックス API 要件セット1.5 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 966c2065268d973ac8476fda839d2a6cdf038f4e
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570738"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="b98a1-103">コンテキスト (メールボックス要件セット 1.5)</span><span class="sxs-lookup"><span data-stu-id="b98a1-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b98a1-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b98a1-104">[Office](office.md).context</span></span>

<span data-ttu-id="b98a1-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b98a1-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b98a1-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b98a1-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-107">Requirements</span></span>

|<span data-ttu-id="b98a1-108">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-108">Requirement</span></span>| <span data-ttu-id="b98a1-109">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-111">1.1</span></span>|
|[<span data-ttu-id="b98a1-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b98a1-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="b98a1-114">Properties</span></span>

| <span data-ttu-id="b98a1-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="b98a1-115">Property</span></span> | <span data-ttu-id="b98a1-116">モード</span><span class="sxs-lookup"><span data-stu-id="b98a1-116">Modes</span></span> | <span data-ttu-id="b98a1-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-117">Return type</span></span> | <span data-ttu-id="b98a1-118">最小値</span><span class="sxs-lookup"><span data-stu-id="b98a1-118">Minimum</span></span><br><span data-ttu-id="b98a1-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="b98a1-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b98a1-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b98a1-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b98a1-121">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-121">Compose</span></span><br><span data-ttu-id="b98a1-122">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-122">Read</span></span> | <span data-ttu-id="b98a1-123">文字列</span><span class="sxs-lookup"><span data-stu-id="b98a1-123">String</span></span> | [<span data-ttu-id="b98a1-124">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-125">ダン</span><span class="sxs-lookup"><span data-stu-id="b98a1-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b98a1-126">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-126">Compose</span></span><br><span data-ttu-id="b98a1-127">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-127">Read</span></span> | [<span data-ttu-id="b98a1-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b98a1-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b98a1-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b98a1-131">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-131">Compose</span></span><br><span data-ttu-id="b98a1-132">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-132">Read</span></span> | <span data-ttu-id="b98a1-133">文字列</span><span class="sxs-lookup"><span data-stu-id="b98a1-133">String</span></span> | [<span data-ttu-id="b98a1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-135">主催</span><span class="sxs-lookup"><span data-stu-id="b98a1-135">host</span></span>](#host-hosttype) | <span data-ttu-id="b98a1-136">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-136">Compose</span></span><br><span data-ttu-id="b98a1-137">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-137">Read</span></span> | [<span data-ttu-id="b98a1-138">HostType</span><span class="sxs-lookup"><span data-stu-id="b98a1-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-139">1.5</span><span class="sxs-lookup"><span data-stu-id="b98a1-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b98a1-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="b98a1-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b98a1-141">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-141">Compose</span></span><br><span data-ttu-id="b98a1-142">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-142">Read</span></span> | [<span data-ttu-id="b98a1-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="b98a1-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-145">platform</span><span class="sxs-lookup"><span data-stu-id="b98a1-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b98a1-146">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-146">Compose</span></span><br><span data-ttu-id="b98a1-147">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-147">Read</span></span> | [<span data-ttu-id="b98a1-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b98a1-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-149">1.5</span><span class="sxs-lookup"><span data-stu-id="b98a1-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b98a1-150">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b98a1-151">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-151">Compose</span></span><br><span data-ttu-id="b98a1-152">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-152">Read</span></span> | [<span data-ttu-id="b98a1-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b98a1-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-154">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b98a1-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b98a1-156">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-156">Compose</span></span><br><span data-ttu-id="b98a1-157">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-157">Read</span></span> | [<span data-ttu-id="b98a1-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b98a1-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b98a1-160">UI</span><span class="sxs-lookup"><span data-stu-id="b98a1-160">ui</span></span>](#ui-ui) | <span data-ttu-id="b98a1-161">作成</span><span class="sxs-lookup"><span data-stu-id="b98a1-161">Compose</span></span><br><span data-ttu-id="b98a1-162">読み取り</span><span class="sxs-lookup"><span data-stu-id="b98a1-162">Read</span></span> | [<span data-ttu-id="b98a1-163">UI</span><span class="sxs-lookup"><span data-stu-id="b98a1-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="b98a1-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b98a1-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="b98a1-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="b98a1-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b98a1-166">contentLanguage: String</span></span>

<span data-ttu-id="b98a1-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b98a1-168">この `contentLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **編集言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-169">型</span><span class="sxs-lookup"><span data-stu-id="b98a1-169">Type</span></span>

*   <span data-ttu-id="b98a1-170">String</span><span class="sxs-lookup"><span data-stu-id="b98a1-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b98a1-171">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-171">Requirements</span></span>

|<span data-ttu-id="b98a1-172">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-172">Requirement</span></span>| <span data-ttu-id="b98a1-173">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-175">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-175">1.1</span></span>|
|[<span data-ttu-id="b98a1-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-178">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b98a1-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b98a1-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b98a1-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-181">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-181">Type</span></span>

*   [<span data-ttu-id="b98a1-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b98a1-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b98a1-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-183">Requirements</span></span>

|<span data-ttu-id="b98a1-184">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-184">Requirement</span></span>| <span data-ttu-id="b98a1-185">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-187">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-187">1.1</span></span>|
|[<span data-ttu-id="b98a1-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-190">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b98a1-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b98a1-191">displayLanguage: String</span></span>

<span data-ttu-id="b98a1-192">Office クライアントアプリケーションの UI 用にユーザーによって指定された RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="b98a1-193">この `displayLanguage` 値は、Office クライアントアプリケーションの [**ファイル > オプション > 言語** で指定されている現在の **表示言語** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-194">型</span><span class="sxs-lookup"><span data-stu-id="b98a1-194">Type</span></span>

*   <span data-ttu-id="b98a1-195">String</span><span class="sxs-lookup"><span data-stu-id="b98a1-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b98a1-196">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-196">Requirements</span></span>

|<span data-ttu-id="b98a1-197">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-197">Requirement</span></span>| <span data-ttu-id="b98a1-198">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-200">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-200">1.1</span></span>|
|[<span data-ttu-id="b98a1-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-203">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b98a1-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b98a1-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b98a1-205">アドインをホストしている Office アプリケーションを取得します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b98a1-206">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してホストを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="b98a1-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-207">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-207">Type</span></span>

*   [<span data-ttu-id="b98a1-208">HostType</span><span class="sxs-lookup"><span data-stu-id="b98a1-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b98a1-209">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-209">Requirements</span></span>

|<span data-ttu-id="b98a1-210">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-210">Requirement</span></span>| <span data-ttu-id="b98a1-211">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-212">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-213">1.5</span><span class="sxs-lookup"><span data-stu-id="b98a1-213">1.5</span></span>|
|[<span data-ttu-id="b98a1-214">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-215">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-216">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b98a1-217">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b98a1-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b98a1-218">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="b98a1-219">別の方法として、 [Office](#diagnostics-contextinformation) のプロパティを使用してプラットフォームを取得することもできます。</span><span class="sxs-lookup"><span data-stu-id="b98a1-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-220">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-220">Type</span></span>

*   [<span data-ttu-id="b98a1-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b98a1-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b98a1-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-222">Requirements</span></span>

|<span data-ttu-id="b98a1-223">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-223">Requirement</span></span>| <span data-ttu-id="b98a1-224">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-225">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-226">1.5</span><span class="sxs-lookup"><span data-stu-id="b98a1-226">1.5</span></span>|
|[<span data-ttu-id="b98a1-227">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-228">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-229">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b98a1-230">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b98a1-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b98a1-231">現在のアプリケーションとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-232">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-232">Type</span></span>

*   [<span data-ttu-id="b98a1-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b98a1-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b98a1-234">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-234">Requirements</span></span>

|<span data-ttu-id="b98a1-235">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-235">Requirement</span></span>| <span data-ttu-id="b98a1-236">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-237">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-238">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-238">1.1</span></span>|
|[<span data-ttu-id="b98a1-239">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-240">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b98a1-241">例</span><span class="sxs-lookup"><span data-stu-id="b98a1-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b98a1-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b98a1-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b98a1-243">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b98a1-244">このオブジェクトを使用する `RoamingSettings` と、ユーザーのメールボックスに格納されているメールアドインのデータを格納してアクセスできます。そのため、そのメールボックスにアクセスするときに使用する Outlook クライアントから実行しているときに、そのアドインが使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="b98a1-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-245">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-245">Type</span></span>

*   [<span data-ttu-id="b98a1-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b98a1-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b98a1-247">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-247">Requirements</span></span>

|<span data-ttu-id="b98a1-248">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-248">Requirement</span></span>| <span data-ttu-id="b98a1-249">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-250">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-251">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-251">1.1</span></span>|
|[<span data-ttu-id="b98a1-252">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b98a1-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b98a1-253">制限あり</span><span class="sxs-lookup"><span data-stu-id="b98a1-253">Restricted</span></span>|
|[<span data-ttu-id="b98a1-254">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-255">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b98a1-256">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b98a1-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b98a1-257">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b98a1-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b98a1-258">種類</span><span class="sxs-lookup"><span data-stu-id="b98a1-258">Type</span></span>

*   [<span data-ttu-id="b98a1-259">UI</span><span class="sxs-lookup"><span data-stu-id="b98a1-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b98a1-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="b98a1-260">Requirements</span></span>

|<span data-ttu-id="b98a1-261">要件</span><span class="sxs-lookup"><span data-stu-id="b98a1-261">Requirement</span></span>| <span data-ttu-id="b98a1-262">値</span><span class="sxs-lookup"><span data-stu-id="b98a1-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="b98a1-263">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b98a1-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b98a1-264">1.1</span><span class="sxs-lookup"><span data-stu-id="b98a1-264">1.1</span></span>|
|[<span data-ttu-id="b98a1-265">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b98a1-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b98a1-266">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b98a1-266">Compose or Read</span></span>|
