---
title: Office コンテキスト要件セット1.3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8cac4f00b67bd49ea4762f9b91dcb7a0df8eeeea
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268664"
---
# <a name="context"></a><span data-ttu-id="7b1b9-102">context</span><span class="sxs-lookup"><span data-stu-id="7b1b9-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="7b1b9-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="7b1b9-103">[Office](Office.md).context</span></span>

<span data-ttu-id="7b1b9-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7b1b9-106">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-106">Requirements</span></span>

|<span data-ttu-id="7b1b9-107">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-107">Requirement</span></span>| <span data-ttu-id="7b1b9-108">値</span><span class="sxs-lookup"><span data-stu-id="7b1b9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7b1b9-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7b1b9-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7b1b9-110">1.0</span><span class="sxs-lookup"><span data-stu-id="7b1b9-110">1.0</span></span>|
|[<span data-ttu-id="7b1b9-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7b1b9-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7b1b9-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7b1b9-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7b1b9-113">メンバーとメソッド</span><span class="sxs-lookup"><span data-stu-id="7b1b9-113">Members and methods</span></span>

| <span data-ttu-id="7b1b9-114">メンバー</span><span class="sxs-lookup"><span data-stu-id="7b1b9-114">Member</span></span> | <span data-ttu-id="7b1b9-115">種類</span><span class="sxs-lookup"><span data-stu-id="7b1b9-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7b1b9-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7b1b9-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7b1b9-117">Member</span><span class="sxs-lookup"><span data-stu-id="7b1b9-117">Member</span></span> |
| [<span data-ttu-id="7b1b9-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7b1b9-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7b1b9-119">メンバー</span><span class="sxs-lookup"><span data-stu-id="7b1b9-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="7b1b9-120">名前空間</span><span class="sxs-lookup"><span data-stu-id="7b1b9-120">Namespaces</span></span>

<span data-ttu-id="7b1b9-121">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="7b1b9-122">Members</span><span class="sxs-lookup"><span data-stu-id="7b1b9-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="7b1b9-123">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7b1b9-123">displayLanguage: String</span></span>

<span data-ttu-id="7b1b9-124">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="7b1b9-125">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="7b1b9-126">型</span><span class="sxs-lookup"><span data-stu-id="7b1b9-126">Type</span></span>

*   <span data-ttu-id="7b1b9-127">String</span><span class="sxs-lookup"><span data-stu-id="7b1b9-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7b1b9-128">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-128">Requirements</span></span>

|<span data-ttu-id="7b1b9-129">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-129">Requirement</span></span>| <span data-ttu-id="7b1b9-130">値</span><span class="sxs-lookup"><span data-stu-id="7b1b9-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="7b1b9-131">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7b1b9-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7b1b9-132">1.0</span><span class="sxs-lookup"><span data-stu-id="7b1b9-132">1.0</span></span>|
|[<span data-ttu-id="7b1b9-133">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7b1b9-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7b1b9-134">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7b1b9-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7b1b9-135">例</span><span class="sxs-lookup"><span data-stu-id="7b1b9-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-13"></a><span data-ttu-id="7b1b9-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="7b1b9-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span></span>

<span data-ttu-id="7b1b9-137">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7b1b9-138">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="7b1b9-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7b1b9-139">型</span><span class="sxs-lookup"><span data-stu-id="7b1b9-139">Type</span></span>

*   [<span data-ttu-id="7b1b9-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7b1b9-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="7b1b9-141">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-141">Requirements</span></span>

|<span data-ttu-id="7b1b9-142">要件</span><span class="sxs-lookup"><span data-stu-id="7b1b9-142">Requirement</span></span>| <span data-ttu-id="7b1b9-143">値</span><span class="sxs-lookup"><span data-stu-id="7b1b9-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="7b1b9-144">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="7b1b9-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7b1b9-145">1.0</span><span class="sxs-lookup"><span data-stu-id="7b1b9-145">1.0</span></span>|
|[<span data-ttu-id="7b1b9-146">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="7b1b9-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7b1b9-147">制限あり</span><span class="sxs-lookup"><span data-stu-id="7b1b9-147">Restricted</span></span>|
|[<span data-ttu-id="7b1b9-148">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="7b1b9-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="7b1b9-149">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="7b1b9-149">Compose or Read</span></span>|
