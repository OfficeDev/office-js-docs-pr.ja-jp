---
title: Office.context - 要件セット 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 8e8d83500e1cb9de730f6a7bac75c3bcc0012b4e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068134"
---
# <a name="context"></a><span data-ttu-id="41cdb-102">context</span><span class="sxs-lookup"><span data-stu-id="41cdb-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="41cdb-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="41cdb-103">[Office](Office.md).context</span></span>

<span data-ttu-id="41cdb-p101">Office.context 名前空間は、すべての Office アプリのアドインで使う共有インターフェイスを提供します。この一覧は、Outlook のアドインで使うインターフェイスのみを記載しています。Office.context 名前空間の完全な一覧は、「[共通 API の Office.context リファレンス](/javascript/api/office/office.context)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="41cdb-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="41cdb-106">要件</span><span class="sxs-lookup"><span data-stu-id="41cdb-106">Requirements</span></span>

|<span data-ttu-id="41cdb-107">要件</span><span class="sxs-lookup"><span data-stu-id="41cdb-107">Requirement</span></span>| <span data-ttu-id="41cdb-108">値</span><span class="sxs-lookup"><span data-stu-id="41cdb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="41cdb-109">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41cdb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="41cdb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="41cdb-110">1.0</span></span>|
|[<span data-ttu-id="41cdb-111">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41cdb-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="41cdb-112">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="41cdb-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="41cdb-113">名前空間</span><span class="sxs-lookup"><span data-stu-id="41cdb-113">Namespaces</span></span>

<span data-ttu-id="41cdb-114">[mailbox](office.context.mailbox.md) - Microsoft Outlook と Microsoft Outlook on the web の Outlook アドイン オブジェクト モデルへのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="41cdb-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="41cdb-115">メンバー</span><span class="sxs-lookup"><span data-stu-id="41cdb-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="41cdb-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="41cdb-116">displayLanguage :String</span></span>

<span data-ttu-id="41cdb-117">Office ホスト アプリケーションの UI 用にユーザーが指定した RFC 1766 言語タグ形式のロケール (言語) を取得します。</span><span class="sxs-lookup"><span data-stu-id="41cdb-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="41cdb-118">`displayLanguage` の値は、Office ホスト アプリケーションの **[ファイル]、[選択肢]、[言語]** によって指定される現在の **[表示言語]** 設定を反映します。</span><span class="sxs-lookup"><span data-stu-id="41cdb-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="41cdb-119">Type</span><span class="sxs-lookup"><span data-stu-id="41cdb-119">Type</span></span>

*   <span data-ttu-id="41cdb-120">String</span><span class="sxs-lookup"><span data-stu-id="41cdb-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="41cdb-121">Requirements</span><span class="sxs-lookup"><span data-stu-id="41cdb-121">Requirements</span></span>

|<span data-ttu-id="41cdb-122">要件</span><span class="sxs-lookup"><span data-stu-id="41cdb-122">Requirement</span></span>| <span data-ttu-id="41cdb-123">値</span><span class="sxs-lookup"><span data-stu-id="41cdb-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="41cdb-124">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41cdb-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="41cdb-125">1.0</span><span class="sxs-lookup"><span data-stu-id="41cdb-125">1.0</span></span>|
|[<span data-ttu-id="41cdb-126">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41cdb-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="41cdb-127">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="41cdb-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="41cdb-128">例</span><span class="sxs-lookup"><span data-stu-id="41cdb-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="41cdb-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="41cdb-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="41cdb-130">ユーザーのメールボックスに保存されている、メール アドインのカスタム設定や状態を表すオブジェクトを取得します。</span><span class="sxs-lookup"><span data-stu-id="41cdb-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="41cdb-131">`RoamingSettings` オブジェクトを使うと、ユーザーのメールボックスに保存されている、メール アドインのデータの保存やアクセスを実行できます。そのため、メール アドインは、このメールボックスへのアクセスに使うどのホスト クライアント アプリケーションから実行されても、このデータを使うことができます。</span><span class="sxs-lookup"><span data-stu-id="41cdb-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="41cdb-132">型</span><span class="sxs-lookup"><span data-stu-id="41cdb-132">Type</span></span>

*   [<span data-ttu-id="41cdb-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="41cdb-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="41cdb-134">要件</span><span class="sxs-lookup"><span data-stu-id="41cdb-134">Requirements</span></span>

|<span data-ttu-id="41cdb-135">要件</span><span class="sxs-lookup"><span data-stu-id="41cdb-135">Requirement</span></span>| <span data-ttu-id="41cdb-136">値</span><span class="sxs-lookup"><span data-stu-id="41cdb-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="41cdb-137">メールボックスの最小要件セットのバージョン</span><span class="sxs-lookup"><span data-stu-id="41cdb-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="41cdb-138">1.0</span><span class="sxs-lookup"><span data-stu-id="41cdb-138">1.0</span></span>|
|[<span data-ttu-id="41cdb-139">最小限のアクセス許可レベル</span><span class="sxs-lookup"><span data-stu-id="41cdb-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="41cdb-140">制限あり</span><span class="sxs-lookup"><span data-stu-id="41cdb-140">Restricted</span></span>|
|[<span data-ttu-id="41cdb-141">適用可能な Outlook のモード</span><span class="sxs-lookup"><span data-stu-id="41cdb-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="41cdb-142">新規作成または読み取り</span><span class="sxs-lookup"><span data-stu-id="41cdb-142">Compose or Read</span></span>|
