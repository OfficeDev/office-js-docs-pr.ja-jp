---
title: Office.context - 要件セット 1.5
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 10e1c9a8b7ba4d62ffb2694cc7cb8edcad15fba7
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064691"
---
# <a name="context"></a><span data-ttu-id="fbbfd-102">context</span><span class="sxs-lookup"><span data-stu-id="fbbfd-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="fbbfd-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="fbbfd-103">[Office](Office.md).context</span></span>

<span data-ttu-id="fbbfd-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbbfd-106">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-106">Requirements</span></span>

|<span data-ttu-id="fbbfd-107">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-107">Requirement</span></span>| <span data-ttu-id="fbbfd-108">値</span><span class="sxs-lookup"><span data-stu-id="fbbfd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbbfd-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fbbfd-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbbfd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="fbbfd-110">1.0</span></span>|
|[<span data-ttu-id="fbbfd-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fbbfd-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fbbfd-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fbbfd-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fbbfd-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="fbbfd-113">Members and methods</span></span>

| <span data-ttu-id="fbbfd-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="fbbfd-114">Member</span></span> | <span data-ttu-id="fbbfd-115">種類</span><span class="sxs-lookup"><span data-stu-id="fbbfd-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fbbfd-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="fbbfd-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="fbbfd-117">Member</span><span class="sxs-lookup"><span data-stu-id="fbbfd-117">Member</span></span> |
| [<span data-ttu-id="fbbfd-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="fbbfd-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="fbbfd-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="fbbfd-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fbbfd-120">名前空間</span><span class="sxs-lookup"><span data-stu-id="fbbfd-120">Namespaces</span></span>

<span data-ttu-id="fbbfd-121">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="fbbfd-122">Members</span><span class="sxs-lookup"><span data-stu-id="fbbfd-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="fbbfd-123">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="fbbfd-123">displayLanguage: String</span></span>

<span data-ttu-id="fbbfd-124">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="fbbfd-125">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="fbbfd-126">型</span><span class="sxs-lookup"><span data-stu-id="fbbfd-126">Type</span></span>

*   <span data-ttu-id="fbbfd-127">String</span><span class="sxs-lookup"><span data-stu-id="fbbfd-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbbfd-128">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-128">Requirements</span></span>

|<span data-ttu-id="fbbfd-129">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-129">Requirement</span></span>| <span data-ttu-id="fbbfd-130">値</span><span class="sxs-lookup"><span data-stu-id="fbbfd-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbbfd-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fbbfd-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbbfd-132">1.0</span><span class="sxs-lookup"><span data-stu-id="fbbfd-132">1.0</span></span>|
|[<span data-ttu-id="fbbfd-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fbbfd-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fbbfd-134">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fbbfd-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbbfd-135">例</span><span class="sxs-lookup"><span data-stu-id="fbbfd-135">Example</span></span>

```javascript
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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-15"></a><span data-ttu-id="fbbfd-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="fbbfd-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)</span></span>

<span data-ttu-id="fbbfd-137">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="fbbfd-138">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="fbbfd-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="fbbfd-139">型</span><span class="sxs-lookup"><span data-stu-id="fbbfd-139">Type</span></span>

*   [<span data-ttu-id="fbbfd-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="fbbfd-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="fbbfd-141">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-141">Requirements</span></span>

|<span data-ttu-id="fbbfd-142">要件</span><span class="sxs-lookup"><span data-stu-id="fbbfd-142">Requirement</span></span>| <span data-ttu-id="fbbfd-143">値</span><span class="sxs-lookup"><span data-stu-id="fbbfd-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbbfd-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="fbbfd-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbbfd-145">1.0</span><span class="sxs-lookup"><span data-stu-id="fbbfd-145">1.0</span></span>|
|[<span data-ttu-id="fbbfd-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="fbbfd-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbbfd-147">制限あり</span><span class="sxs-lookup"><span data-stu-id="fbbfd-147">Restricted</span></span>|
|[<span data-ttu-id="fbbfd-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="fbbfd-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fbbfd-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="fbbfd-149">Compose or Read</span></span>|
