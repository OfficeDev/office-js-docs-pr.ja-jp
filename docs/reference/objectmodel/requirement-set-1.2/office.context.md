---
title: Office コンテキスト要件セット1.2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 194532acc9262357a54e355fcfa9b15336cd95b0
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064292"
---
# <a name="context"></a><span data-ttu-id="1e873-102">context</span><span class="sxs-lookup"><span data-stu-id="1e873-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="1e873-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="1e873-103">[Office](Office.md).context</span></span>

<span data-ttu-id="1e873-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1e873-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="1e873-106">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-106">Requirements</span></span>

|<span data-ttu-id="1e873-107">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-107">Requirement</span></span>| <span data-ttu-id="1e873-108">値</span><span class="sxs-lookup"><span data-stu-id="1e873-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e873-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e873-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e873-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1e873-110">1.0</span></span>|
|[<span data-ttu-id="1e873-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e873-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e873-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e873-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="1e873-113">名前空間</span><span class="sxs-lookup"><span data-stu-id="1e873-113">Namespaces</span></span>

<span data-ttu-id="1e873-114">[mailbox](office.context.mailbox.md) -Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="1e873-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="1e873-115">Members</span><span class="sxs-lookup"><span data-stu-id="1e873-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="1e873-116">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1e873-116">displayLanguage: String</span></span>

<span data-ttu-id="1e873-117">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="1e873-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1e873-118">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="1e873-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1e873-119">型</span><span class="sxs-lookup"><span data-stu-id="1e873-119">Type</span></span>

*   <span data-ttu-id="1e873-120">String</span><span class="sxs-lookup"><span data-stu-id="1e873-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1e873-121">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-121">Requirements</span></span>

|<span data-ttu-id="1e873-122">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-122">Requirement</span></span>| <span data-ttu-id="1e873-123">値</span><span class="sxs-lookup"><span data-stu-id="1e873-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e873-124">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e873-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e873-125">1.0</span><span class="sxs-lookup"><span data-stu-id="1e873-125">1.0</span></span>|
|[<span data-ttu-id="1e873-126">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e873-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e873-127">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e873-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1e873-128">例</span><span class="sxs-lookup"><span data-stu-id="1e873-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-12"></a><span data-ttu-id="1e873-129">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1e873-129">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1e873-130">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="1e873-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1e873-131">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="1e873-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1e873-132">型</span><span class="sxs-lookup"><span data-stu-id="1e873-132">Type</span></span>

*   [<span data-ttu-id="1e873-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1e873-133">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1e873-134">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-134">Requirements</span></span>

|<span data-ttu-id="1e873-135">要件</span><span class="sxs-lookup"><span data-stu-id="1e873-135">Requirement</span></span>| <span data-ttu-id="1e873-136">値</span><span class="sxs-lookup"><span data-stu-id="1e873-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="1e873-137">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="1e873-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1e873-138">1.0</span><span class="sxs-lookup"><span data-stu-id="1e873-138">1.0</span></span>|
|[<span data-ttu-id="1e873-139">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="1e873-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1e873-140">制限あり</span><span class="sxs-lookup"><span data-stu-id="1e873-140">Restricted</span></span>|
|[<span data-ttu-id="1e873-141">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="1e873-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1e873-142">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="1e873-142">Compose or Read</span></span>|
