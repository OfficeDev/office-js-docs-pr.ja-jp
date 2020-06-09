---
title: Office コンテキスト要件セット1.2
description: メールボックス API 要件セット1.2 を使用した Outlook アドインで使用可能な Office コンテキストオブジェクトメンバー。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 4bff6e70b143480b0d4b98925f504870d8a2bbb2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610527"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="b926a-103">コンテキスト (メールボックス要件セット 1.2)</span><span class="sxs-lookup"><span data-stu-id="b926a-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b926a-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b926a-104">[Office](office.md).context</span></span>

<span data-ttu-id="b926a-105">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="b926a-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b926a-106">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.2)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b926a-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b926a-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-107">Requirements</span></span>

|<span data-ttu-id="b926a-108">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-108">Requirement</span></span>| <span data-ttu-id="b926a-109">値</span><span class="sxs-lookup"><span data-stu-id="b926a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-110">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-111">1.1</span></span>|
|[<span data-ttu-id="b926a-112">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-113">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b926a-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="b926a-114">Properties</span></span>

| <span data-ttu-id="b926a-115">プロパティ</span><span class="sxs-lookup"><span data-stu-id="b926a-115">Property</span></span> | <span data-ttu-id="b926a-116">モード</span><span class="sxs-lookup"><span data-stu-id="b926a-116">Modes</span></span> | <span data-ttu-id="b926a-117">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="b926a-117">Return type</span></span> | <span data-ttu-id="b926a-118">最小値</span><span class="sxs-lookup"><span data-stu-id="b926a-118">Minimum</span></span><br><span data-ttu-id="b926a-119">要件セット</span><span class="sxs-lookup"><span data-stu-id="b926a-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b926a-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b926a-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b926a-121">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-121">Compose</span></span><br><span data-ttu-id="b926a-122">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-122">Read</span></span> | <span data-ttu-id="b926a-123">String</span><span class="sxs-lookup"><span data-stu-id="b926a-123">String</span></span> | [<span data-ttu-id="b926a-124">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-125">ダン</span><span class="sxs-lookup"><span data-stu-id="b926a-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b926a-126">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-126">Compose</span></span><br><span data-ttu-id="b926a-127">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-127">Read</span></span> | [<span data-ttu-id="b926a-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b926a-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [<span data-ttu-id="b926a-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b926a-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b926a-131">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-131">Compose</span></span><br><span data-ttu-id="b926a-132">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-132">Read</span></span> | <span data-ttu-id="b926a-133">String</span><span class="sxs-lookup"><span data-stu-id="b926a-133">String</span></span> | [<span data-ttu-id="b926a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-135">主催</span><span class="sxs-lookup"><span data-stu-id="b926a-135">host</span></span>](#host-hosttype) | <span data-ttu-id="b926a-136">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-136">Compose</span></span><br><span data-ttu-id="b926a-137">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-137">Read</span></span> | [<span data-ttu-id="b926a-138">HostType</span><span class="sxs-lookup"><span data-stu-id="b926a-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [<span data-ttu-id="b926a-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="b926a-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b926a-141">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-141">Compose</span></span><br><span data-ttu-id="b926a-142">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-142">Read</span></span> | [<span data-ttu-id="b926a-143">メールボックス</span><span class="sxs-lookup"><span data-stu-id="b926a-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2) | [<span data-ttu-id="b926a-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-145">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="b926a-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b926a-146">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-146">Compose</span></span><br><span data-ttu-id="b926a-147">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-147">Read</span></span> | [<span data-ttu-id="b926a-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b926a-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [<span data-ttu-id="b926a-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-150">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b926a-151">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-151">Compose</span></span><br><span data-ttu-id="b926a-152">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-152">Read</span></span> | [<span data-ttu-id="b926a-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b926a-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [<span data-ttu-id="b926a-154">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b926a-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b926a-156">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-156">Compose</span></span><br><span data-ttu-id="b926a-157">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-157">Read</span></span> | [<span data-ttu-id="b926a-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b926a-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [<span data-ttu-id="b926a-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b926a-160">UI</span><span class="sxs-lookup"><span data-stu-id="b926a-160">ui</span></span>](#ui-ui) | <span data-ttu-id="b926a-161">作成</span><span class="sxs-lookup"><span data-stu-id="b926a-161">Compose</span></span><br><span data-ttu-id="b926a-162">Read</span><span class="sxs-lookup"><span data-stu-id="b926a-162">Read</span></span> | [<span data-ttu-id="b926a-163">UI</span><span class="sxs-lookup"><span data-stu-id="b926a-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2) | [<span data-ttu-id="b926a-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b926a-165">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="b926a-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="b926a-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b926a-166">contentLanguage: String</span></span>

<span data-ttu-id="b926a-167">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="b926a-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b926a-168">この `contentLanguage` 値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="b926a-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-169">型</span><span class="sxs-lookup"><span data-stu-id="b926a-169">Type</span></span>

*   <span data-ttu-id="b926a-170">String</span><span class="sxs-lookup"><span data-stu-id="b926a-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b926a-171">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-171">Requirements</span></span>

|<span data-ttu-id="b926a-172">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-172">Requirement</span></span>| <span data-ttu-id="b926a-173">値</span><span class="sxs-lookup"><span data-stu-id="b926a-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-174">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-175">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-175">1.1</span></span>|
|[<span data-ttu-id="b926a-176">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-177">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-178">例</span><span class="sxs-lookup"><span data-stu-id="b926a-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b926a-179">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b926a-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b926a-180">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="b926a-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-181">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-181">Type</span></span>

*   [<span data-ttu-id="b926a-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b926a-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b926a-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-183">Requirements</span></span>

|<span data-ttu-id="b926a-184">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-184">Requirement</span></span>| <span data-ttu-id="b926a-185">値</span><span class="sxs-lookup"><span data-stu-id="b926a-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-186">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-187">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-187">1.1</span></span>|
|[<span data-ttu-id="b926a-188">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-189">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-190">例</span><span class="sxs-lookup"><span data-stu-id="b926a-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b926a-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b926a-191">displayLanguage: String</span></span>

<span data-ttu-id="b926a-192">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="b926a-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b926a-193">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="b926a-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-194">型</span><span class="sxs-lookup"><span data-stu-id="b926a-194">Type</span></span>

*   <span data-ttu-id="b926a-195">String</span><span class="sxs-lookup"><span data-stu-id="b926a-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b926a-196">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-196">Requirements</span></span>

|<span data-ttu-id="b926a-197">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-197">Requirement</span></span>| <span data-ttu-id="b926a-198">値</span><span class="sxs-lookup"><span data-stu-id="b926a-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-199">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-200">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-200">1.1</span></span>|
|[<span data-ttu-id="b926a-201">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-202">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-203">例</span><span class="sxs-lookup"><span data-stu-id="b926a-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b926a-204">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b926a-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b926a-205">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="b926a-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-206">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-206">Type</span></span>

*   [<span data-ttu-id="b926a-207">HostType</span><span class="sxs-lookup"><span data-stu-id="b926a-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b926a-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-208">Requirements</span></span>

|<span data-ttu-id="b926a-209">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-209">Requirement</span></span>| <span data-ttu-id="b926a-210">値</span><span class="sxs-lookup"><span data-stu-id="b926a-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-211">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-212">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-212">1.1</span></span>|
|[<span data-ttu-id="b926a-213">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-214">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-215">例</span><span class="sxs-lookup"><span data-stu-id="b926a-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b926a-216">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b926a-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b926a-217">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="b926a-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-218">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-218">Type</span></span>

*   [<span data-ttu-id="b926a-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b926a-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b926a-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-220">Requirements</span></span>

|<span data-ttu-id="b926a-221">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-221">Requirement</span></span>| <span data-ttu-id="b926a-222">値</span><span class="sxs-lookup"><span data-stu-id="b926a-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-223">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-224">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-224">1.1</span></span>|
|[<span data-ttu-id="b926a-225">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-226">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-227">例</span><span class="sxs-lookup"><span data-stu-id="b926a-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b926a-228">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b926a-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b926a-229">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b926a-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-230">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-230">Type</span></span>

*   [<span data-ttu-id="b926a-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b926a-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b926a-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-232">Requirements</span></span>

|<span data-ttu-id="b926a-233">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-233">Requirement</span></span>| <span data-ttu-id="b926a-234">値</span><span class="sxs-lookup"><span data-stu-id="b926a-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-235">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-236">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-236">1.1</span></span>|
|[<span data-ttu-id="b926a-237">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-238">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b926a-239">例</span><span class="sxs-lookup"><span data-stu-id="b926a-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b926a-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b926a-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b926a-241">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="b926a-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b926a-242">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="b926a-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-243">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-243">Type</span></span>

*   [<span data-ttu-id="b926a-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b926a-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b926a-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-245">Requirements</span></span>

|<span data-ttu-id="b926a-246">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-246">Requirement</span></span>| <span data-ttu-id="b926a-247">値</span><span class="sxs-lookup"><span data-stu-id="b926a-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-248">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-249">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-249">1.1</span></span>|
|[<span data-ttu-id="b926a-250">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="b926a-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b926a-251">制限あり</span><span class="sxs-lookup"><span data-stu-id="b926a-251">Restricted</span></span>|
|[<span data-ttu-id="b926a-252">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-253">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b926a-254">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b926a-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b926a-255">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="b926a-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b926a-256">種類</span><span class="sxs-lookup"><span data-stu-id="b926a-256">Type</span></span>

*   [<span data-ttu-id="b926a-257">UI</span><span class="sxs-lookup"><span data-stu-id="b926a-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b926a-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="b926a-258">Requirements</span></span>

|<span data-ttu-id="b926a-259">要件</span><span class="sxs-lookup"><span data-stu-id="b926a-259">Requirement</span></span>| <span data-ttu-id="b926a-260">値</span><span class="sxs-lookup"><span data-stu-id="b926a-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="b926a-261">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="b926a-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b926a-262">1.1</span><span class="sxs-lookup"><span data-stu-id="b926a-262">1.1</span></span>|
|[<span data-ttu-id="b926a-263">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="b926a-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b926a-264">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="b926a-264">Compose or Read</span></span>|
