---
title: Office コンテキスト要件セット1.4
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 7f4637a1d6a4a9bc2f97d039ed4404ab549a2b34
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064649"
---
# <a name="context"></a><span data-ttu-id="60203-102">context</span><span class="sxs-lookup"><span data-stu-id="60203-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="60203-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="60203-103">[Office](Office.md).context</span></span>

<span data-ttu-id="60203-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="60203-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="60203-106">要件</span><span class="sxs-lookup"><span data-stu-id="60203-106">Requirements</span></span>

|<span data-ttu-id="60203-107">要件</span><span class="sxs-lookup"><span data-stu-id="60203-107">Requirement</span></span>| <span data-ttu-id="60203-108">値</span><span class="sxs-lookup"><span data-stu-id="60203-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="60203-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60203-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60203-110">1.0</span><span class="sxs-lookup"><span data-stu-id="60203-110">1.0</span></span>|
|[<span data-ttu-id="60203-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60203-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60203-112">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="60203-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="60203-113">名前空間</span><span class="sxs-lookup"><span data-stu-id="60203-113">Namespaces</span></span>

<span data-ttu-id="60203-114">[メールボックス](office.context.mailbox.md): Microsoft Outlook の outlook アドインオブジェクトモデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="60203-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="60203-115">Members</span><span class="sxs-lookup"><span data-stu-id="60203-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="60203-116">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="60203-116">displayLanguage: String</span></span>

<span data-ttu-id="60203-117">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="60203-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="60203-118">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="60203-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="60203-119">型</span><span class="sxs-lookup"><span data-stu-id="60203-119">Type</span></span>

*   <span data-ttu-id="60203-120">String</span><span class="sxs-lookup"><span data-stu-id="60203-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="60203-121">要件</span><span class="sxs-lookup"><span data-stu-id="60203-121">Requirements</span></span>

|<span data-ttu-id="60203-122">要件</span><span class="sxs-lookup"><span data-stu-id="60203-122">Requirement</span></span>| <span data-ttu-id="60203-123">値</span><span class="sxs-lookup"><span data-stu-id="60203-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="60203-124">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60203-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60203-125">1.0</span><span class="sxs-lookup"><span data-stu-id="60203-125">1.0</span></span>|
|[<span data-ttu-id="60203-126">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60203-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60203-127">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="60203-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="60203-128">例</span><span class="sxs-lookup"><span data-stu-id="60203-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-14"></a><span data-ttu-id="60203-129">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="60203-129">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)</span></span>

<span data-ttu-id="60203-130">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="60203-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="60203-131">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="60203-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="60203-132">型</span><span class="sxs-lookup"><span data-stu-id="60203-132">Type</span></span>

*   [<span data-ttu-id="60203-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="60203-133">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="60203-134">要件</span><span class="sxs-lookup"><span data-stu-id="60203-134">Requirements</span></span>

|<span data-ttu-id="60203-135">要件</span><span class="sxs-lookup"><span data-stu-id="60203-135">Requirement</span></span>| <span data-ttu-id="60203-136">値</span><span class="sxs-lookup"><span data-stu-id="60203-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="60203-137">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="60203-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="60203-138">1.0</span><span class="sxs-lookup"><span data-stu-id="60203-138">1.0</span></span>|
|[<span data-ttu-id="60203-139">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="60203-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="60203-140">制限あり</span><span class="sxs-lookup"><span data-stu-id="60203-140">Restricted</span></span>|
|[<span data-ttu-id="60203-141">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="60203-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="60203-142">新規作成または閲覧</span><span class="sxs-lookup"><span data-stu-id="60203-142">Compose or Read</span></span>|
