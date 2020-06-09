---
title: Office コンテキスト要件セット1.8
description: メールボックス API 要件セット1.8 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 00e9c7fe5b92174f58d97fe75b03e954c3b0c7e9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612186"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="3c4f0-103">コンテキスト (メールボックス要件セット 1.8)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="3c4f0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="3c4f0-104">[Office](office.md).context</span></span>

<span data-ttu-id="3c4f0-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="3c4f0-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.8)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c4f0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-107">Requirements</span></span>

|<span data-ttu-id="3c4f0-108">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-108">Requirement</span></span>| <span data-ttu-id="3c4f0-109">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-111">1.1</span></span>|
|[<span data-ttu-id="3c4f0-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3c4f0-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3c4f0-114">Properties</span></span>

| <span data-ttu-id="3c4f0-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="3c4f0-115">Property</span></span> | <span data-ttu-id="3c4f0-116">モード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-116">Modes</span></span> | <span data-ttu-id="3c4f0-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-117">Return type</span></span> | <span data-ttu-id="3c4f0-118">最小値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-118">Minimum</span></span><br><span data-ttu-id="3c4f0-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="3c4f0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3c4f0-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="3c4f0-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="3c4f0-121">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-121">Compose</span></span><br><span data-ttu-id="3c4f0-122">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-122">Read</span></span> | <span data-ttu-id="3c4f0-123">String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-123">String</span></span> | [<span data-ttu-id="3c4f0-124">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-125">ダン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="3c4f0-126">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-126">Compose</span></span><br><span data-ttu-id="3c4f0-127">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-127">Read</span></span> | [<span data-ttu-id="3c4f0-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3c4f0-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="3c4f0-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="3c4f0-131">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-131">Compose</span></span><br><span data-ttu-id="3c4f0-132">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-132">Read</span></span> | <span data-ttu-id="3c4f0-133">String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-133">String</span></span> | [<span data-ttu-id="3c4f0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-135">主催</span><span class="sxs-lookup"><span data-stu-id="3c4f0-135">host</span></span>](#host-hosttype) | <span data-ttu-id="3c4f0-136">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-136">Compose</span></span><br><span data-ttu-id="3c4f0-137">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-137">Read</span></span> | [<span data-ttu-id="3c4f0-138">HostType</span><span class="sxs-lookup"><span data-stu-id="3c4f0-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="3c4f0-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="3c4f0-141">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-141">Compose</span></span><br><span data-ttu-id="3c4f0-142">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-142">Read</span></span> | [<span data-ttu-id="3c4f0-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="3c4f0-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="3c4f0-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="3c4f0-146">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-146">Compose</span></span><br><span data-ttu-id="3c4f0-147">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-147">Read</span></span> | [<span data-ttu-id="3c4f0-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3c4f0-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-150">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="3c4f0-151">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-151">Compose</span></span><br><span data-ttu-id="3c4f0-152">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-152">Read</span></span> | [<span data-ttu-id="3c4f0-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3c4f0-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-154">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="3c4f0-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="3c4f0-156">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-156">Compose</span></span><br><span data-ttu-id="3c4f0-157">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-157">Read</span></span> | [<span data-ttu-id="3c4f0-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3c4f0-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3c4f0-160">UI</span><span class="sxs-lookup"><span data-stu-id="3c4f0-160">ui</span></span>](#ui-ui) | <span data-ttu-id="3c4f0-161">作成</span><span class="sxs-lookup"><span data-stu-id="3c4f0-161">Compose</span></span><br><span data-ttu-id="3c4f0-162">Read</span><span class="sxs-lookup"><span data-stu-id="3c4f0-162">Read</span></span> | [<span data-ttu-id="3c4f0-163">UI</span><span class="sxs-lookup"><span data-stu-id="3c4f0-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8) | [<span data-ttu-id="3c4f0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="3c4f0-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="3c4f0-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="3c4f0-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-166">contentLanguage: String</span></span>

<span data-ttu-id="3c4f0-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="3c4f0-168">この `contentLanguage` 値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-169">型</span><span class="sxs-lookup"><span data-stu-id="3c4f0-169">Type</span></span>

*   <span data-ttu-id="3c4f0-170">String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c4f0-171">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-171">Requirements</span></span>

|<span data-ttu-id="3c4f0-172">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-172">Requirement</span></span>| <span data-ttu-id="3c4f0-173">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-175">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-175">1.1</span></span>|
|[<span data-ttu-id="3c4f0-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-178">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="3c4f0-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="3c4f0-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-181">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-181">Type</span></span>

*   [<span data-ttu-id="3c4f0-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="3c4f0-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-183">Requirements</span></span>

|<span data-ttu-id="3c4f0-184">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-184">Requirement</span></span>| <span data-ttu-id="3c4f0-185">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-187">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-187">1.1</span></span>|
|[<span data-ttu-id="3c4f0-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-190">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="3c4f0-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-191">displayLanguage: String</span></span>

<span data-ttu-id="3c4f0-192">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="3c4f0-193">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-194">型</span><span class="sxs-lookup"><span data-stu-id="3c4f0-194">Type</span></span>

*   <span data-ttu-id="3c4f0-195">String</span><span class="sxs-lookup"><span data-stu-id="3c4f0-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c4f0-196">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-196">Requirements</span></span>

|<span data-ttu-id="3c4f0-197">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-197">Requirement</span></span>| <span data-ttu-id="3c4f0-198">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-200">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-200">1.1</span></span>|
|[<span data-ttu-id="3c4f0-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-203">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="3c4f0-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="3c4f0-205">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-206">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-206">Type</span></span>

*   [<span data-ttu-id="3c4f0-207">HostType</span><span class="sxs-lookup"><span data-stu-id="3c4f0-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-208">Requirements</span></span>

|<span data-ttu-id="3c4f0-209">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-209">Requirement</span></span>| <span data-ttu-id="3c4f0-210">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-212">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-212">1.1</span></span>|
|[<span data-ttu-id="3c4f0-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-215">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="3c4f0-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="3c4f0-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-218">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-218">Type</span></span>

*   [<span data-ttu-id="3c4f0-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="3c4f0-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-220">Requirements</span></span>

|<span data-ttu-id="3c4f0-221">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-221">Requirement</span></span>| <span data-ttu-id="3c4f0-222">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-224">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-224">1.1</span></span>|
|[<span data-ttu-id="3c4f0-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-227">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="3c4f0-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="3c4f0-229">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-230">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-230">Type</span></span>

*   [<span data-ttu-id="3c4f0-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="3c4f0-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-232">Requirements</span></span>

|<span data-ttu-id="3c4f0-233">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-233">Requirement</span></span>| <span data-ttu-id="3c4f0-234">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-236">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-236">1.1</span></span>|
|[<span data-ttu-id="3c4f0-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c4f0-239">例</span><span class="sxs-lookup"><span data-stu-id="3c4f0-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="3c4f0-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="3c4f0-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3c4f0-242">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-243">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-243">Type</span></span>

*   [<span data-ttu-id="3c4f0-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3c4f0-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-245">Requirements</span></span>

|<span data-ttu-id="3c4f0-246">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-246">Requirement</span></span>| <span data-ttu-id="3c4f0-247">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-249">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-249">1.1</span></span>|
|[<span data-ttu-id="3c4f0-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="3c4f0-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="3c4f0-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="3c4f0-251">Restricted</span></span>|
|[<span data-ttu-id="3c4f0-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="3c4f0-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="3c4f0-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="3c4f0-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="3c4f0-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3c4f0-256">種類</span><span class="sxs-lookup"><span data-stu-id="3c4f0-256">Type</span></span>

*   [<span data-ttu-id="3c4f0-257">UI</span><span class="sxs-lookup"><span data-stu-id="3c4f0-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="3c4f0-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c4f0-258">Requirements</span></span>

|<span data-ttu-id="3c4f0-259">要件</span><span class="sxs-lookup"><span data-stu-id="3c4f0-259">Requirement</span></span>| <span data-ttu-id="3c4f0-260">値</span><span class="sxs-lookup"><span data-stu-id="3c4f0-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c4f0-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="3c4f0-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3c4f0-262">1.1</span><span class="sxs-lookup"><span data-stu-id="3c4f0-262">1.1</span></span>|
|[<span data-ttu-id="3c4f0-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="3c4f0-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3c4f0-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="3c4f0-264">Compose or Read</span></span>|
