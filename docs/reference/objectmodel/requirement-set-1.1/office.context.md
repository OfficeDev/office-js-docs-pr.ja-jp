---
title: Office コンテキスト要件セット1.1
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 973c1604a08e897780f3a12150adffb02eb97b55
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111199"
---
# <a name="context"></a><span data-ttu-id="bec1c-102">context</span><span class="sxs-lookup"><span data-stu-id="bec1c-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="bec1c-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="bec1c-103">[Office](office.md).context</span></span>

<span data-ttu-id="bec1c-104">Office のコンテキストは、すべての Office アプリでアドインによって使用される共有インターフェイスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="bec1c-105">この一覧には、Outlook アドインで使用されるインターフェイスのみが記載されています。Office コンテキスト名前空間の完全な一覧については、 [COMMON API の「office コンテキスト](/javascript/api/office/office.context?view=outlook-js-1.1)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bec1c-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bec1c-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-106">Requirements</span></span>

|<span data-ttu-id="bec1c-107">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-107">Requirement</span></span>| <span data-ttu-id="bec1c-108">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-110">1.1</span></span>|
|[<span data-ttu-id="bec1c-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="bec1c-113">プロパティ</span><span class="sxs-lookup"><span data-stu-id="bec1c-113">Properties</span></span>

| <span data-ttu-id="bec1c-114">プロパティ</span><span class="sxs-lookup"><span data-stu-id="bec1c-114">Property</span></span> | <span data-ttu-id="bec1c-115">モード</span><span class="sxs-lookup"><span data-stu-id="bec1c-115">Modes</span></span> | <span data-ttu-id="bec1c-116">戻り値の種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-116">Return type</span></span> | <span data-ttu-id="bec1c-117">最小値</span><span class="sxs-lookup"><span data-stu-id="bec1c-117">Minimum</span></span><br><span data-ttu-id="bec1c-118">要件セット</span><span class="sxs-lookup"><span data-stu-id="bec1c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bec1c-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="bec1c-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="bec1c-120">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-120">Compose</span></span><br><span data-ttu-id="bec1c-121">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-121">Read</span></span> | <span data-ttu-id="bec1c-122">String</span><span class="sxs-lookup"><span data-stu-id="bec1c-122">String</span></span> | [<span data-ttu-id="bec1c-123">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-124">ダン</span><span class="sxs-lookup"><span data-stu-id="bec1c-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="bec1c-125">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-125">Compose</span></span><br><span data-ttu-id="bec1c-126">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-126">Read</span></span> | [<span data-ttu-id="bec1c-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="bec1c-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-128">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="bec1c-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="bec1c-130">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-130">Compose</span></span><br><span data-ttu-id="bec1c-131">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-131">Read</span></span> | <span data-ttu-id="bec1c-132">String</span><span class="sxs-lookup"><span data-stu-id="bec1c-132">String</span></span> | [<span data-ttu-id="bec1c-133">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-134">主催</span><span class="sxs-lookup"><span data-stu-id="bec1c-134">host</span></span>](#host-hosttype) | <span data-ttu-id="bec1c-135">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-135">Compose</span></span><br><span data-ttu-id="bec1c-136">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-136">Read</span></span> | [<span data-ttu-id="bec1c-137">HostType</span><span class="sxs-lookup"><span data-stu-id="bec1c-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-138">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="bec1c-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="bec1c-140">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-140">Compose</span></span><br><span data-ttu-id="bec1c-141">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-141">Read</span></span> | [<span data-ttu-id="bec1c-142">メールボックス</span><span class="sxs-lookup"><span data-stu-id="bec1c-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-143">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-144">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bec1c-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="bec1c-145">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-145">Compose</span></span><br><span data-ttu-id="bec1c-146">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-146">Read</span></span> | [<span data-ttu-id="bec1c-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="bec1c-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-148">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-149">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="bec1c-150">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-150">Compose</span></span><br><span data-ttu-id="bec1c-151">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-151">Read</span></span> | [<span data-ttu-id="bec1c-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="bec1c-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-153">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="bec1c-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="bec1c-155">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-155">Compose</span></span><br><span data-ttu-id="bec1c-156">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-156">Read</span></span> | [<span data-ttu-id="bec1c-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bec1c-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-158">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bec1c-159">UI</span><span class="sxs-lookup"><span data-stu-id="bec1c-159">ui</span></span>](#ui-ui) | <span data-ttu-id="bec1c-160">作成</span><span class="sxs-lookup"><span data-stu-id="bec1c-160">Compose</span></span><br><span data-ttu-id="bec1c-161">読み取り</span><span class="sxs-lookup"><span data-stu-id="bec1c-161">Read</span></span> | [<span data-ttu-id="bec1c-162">UI</span><span class="sxs-lookup"><span data-stu-id="bec1c-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1) | [<span data-ttu-id="bec1c-163">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="bec1c-164">プロパティの詳細</span><span class="sxs-lookup"><span data-stu-id="bec1c-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="bec1c-165">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="bec1c-165">contentLanguage: String</span></span>

<span data-ttu-id="bec1c-166">アイテムを編集するためにユーザーによって指定されたロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="bec1c-167">この`contentLanguage`値は、Office ホストアプリケーションの [**ファイル > オプション > 言語**で指定されている現在の**編集言語**設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-168">型</span><span class="sxs-lookup"><span data-stu-id="bec1c-168">Type</span></span>

*   <span data-ttu-id="bec1c-169">String</span><span class="sxs-lookup"><span data-stu-id="bec1c-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bec1c-170">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-170">Requirements</span></span>

|<span data-ttu-id="bec1c-171">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-171">Requirement</span></span>| <span data-ttu-id="bec1c-172">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-173">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-174">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-174">1.1</span></span>|
|[<span data-ttu-id="bec1c-175">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-176">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-177">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="bec1c-178">診断: [Contextinformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="bec1c-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="bec1c-179">アドインが実行されている環境に関する情報を取得します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-180">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-180">Type</span></span>

*   [<span data-ttu-id="bec1c-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="bec1c-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="bec1c-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-182">Requirements</span></span>

|<span data-ttu-id="bec1c-183">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-183">Requirement</span></span>| <span data-ttu-id="bec1c-184">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-185">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-186">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-186">1.1</span></span>|
|[<span data-ttu-id="bec1c-187">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-188">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-189">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="bec1c-190">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="bec1c-190">displayLanguage: String</span></span>

<span data-ttu-id="bec1c-191">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="bec1c-192">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-193">型</span><span class="sxs-lookup"><span data-stu-id="bec1c-193">Type</span></span>

*   <span data-ttu-id="bec1c-194">String</span><span class="sxs-lookup"><span data-stu-id="bec1c-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bec1c-195">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-195">Requirements</span></span>

|<span data-ttu-id="bec1c-196">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-196">Requirement</span></span>| <span data-ttu-id="bec1c-197">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-198">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-199">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-199">1.1</span></span>|
|[<span data-ttu-id="bec1c-200">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-201">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-202">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="bec1c-203">ホスト: [Hosttype](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="bec1c-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="bec1c-204">アドインが実行されている Office アプリケーションホストを取得します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-205">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-205">Type</span></span>

*   [<span data-ttu-id="bec1c-206">HostType</span><span class="sxs-lookup"><span data-stu-id="bec1c-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="bec1c-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-207">Requirements</span></span>

|<span data-ttu-id="bec1c-208">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-208">Requirement</span></span>| <span data-ttu-id="bec1c-209">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-210">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-211">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-211">1.1</span></span>|
|[<span data-ttu-id="bec1c-212">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-213">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-214">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="bec1c-215">プラットフォーム: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="bec1c-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="bec1c-216">アドインが実行されているプラットフォームを提供します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-217">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-217">Type</span></span>

*   [<span data-ttu-id="bec1c-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="bec1c-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="bec1c-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-219">Requirements</span></span>

|<span data-ttu-id="bec1c-220">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-220">Requirement</span></span>| <span data-ttu-id="bec1c-221">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-222">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-223">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-223">1.1</span></span>|
|[<span data-ttu-id="bec1c-224">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-225">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-226">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="bec1c-227">要件: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="bec1c-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="bec1c-228">現在のホストとプラットフォームでサポートされている要件セットを判断するためのメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-229">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-229">Type</span></span>

*   [<span data-ttu-id="bec1c-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="bec1c-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="bec1c-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-231">Requirements</span></span>

|<span data-ttu-id="bec1c-232">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-232">Requirement</span></span>| <span data-ttu-id="bec1c-233">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-234">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-235">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-235">1.1</span></span>|
|[<span data-ttu-id="bec1c-236">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-237">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bec1c-238">例</span><span class="sxs-lookup"><span data-stu-id="bec1c-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="bec1c-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="bec1c-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="bec1c-240">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="bec1c-241">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="bec1c-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-242">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-242">Type</span></span>

*   [<span data-ttu-id="bec1c-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bec1c-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="bec1c-244">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-244">Requirements</span></span>

|<span data-ttu-id="bec1c-245">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-245">Requirement</span></span>| <span data-ttu-id="bec1c-246">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-247">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-248">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-248">1.1</span></span>|
|[<span data-ttu-id="bec1c-249">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="bec1c-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bec1c-250">制限あり</span><span class="sxs-lookup"><span data-stu-id="bec1c-250">Restricted</span></span>|
|[<span data-ttu-id="bec1c-251">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-252">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="bec1c-253">ui: [ui](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="bec1c-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="bec1c-254">Office アドインで、ダイアログボックスなどの UI コンポーネントを作成および操作するために使用できるオブジェクトとメソッドを提供します。</span><span class="sxs-lookup"><span data-stu-id="bec1c-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="bec1c-255">種類</span><span class="sxs-lookup"><span data-stu-id="bec1c-255">Type</span></span>

*   [<span data-ttu-id="bec1c-256">UI</span><span class="sxs-lookup"><span data-stu-id="bec1c-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="bec1c-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="bec1c-257">Requirements</span></span>

|<span data-ttu-id="bec1c-258">要件</span><span class="sxs-lookup"><span data-stu-id="bec1c-258">Requirement</span></span>| <span data-ttu-id="bec1c-259">値</span><span class="sxs-lookup"><span data-stu-id="bec1c-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="bec1c-260">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="bec1c-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bec1c-261">1.1</span><span class="sxs-lookup"><span data-stu-id="bec1c-261">1.1</span></span>|
|[<span data-ttu-id="bec1c-262">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="bec1c-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bec1c-263">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="bec1c-263">Compose or Read</span></span>|
