---
title: Office アドインを既存の COM アドインと互換できるようにする
description: Office アドインと同等の COM アドインの互換性を有効にする
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: cfe39dcecafc23bd0813251f7c883aa7a8810273
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718903"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="de518-103">Office アドインを既存の COM アドインと互換できるようにする</span><span class="sxs-lookup"><span data-stu-id="de518-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="de518-104">既存の COM アドインがある場合は、Office アドインで同等の機能を構築できます。これにより、web や Office on the Mac 上の他のプラットフォーム上でソリューションを実行することが可能になります。</span><span class="sxs-lookup"><span data-stu-id="de518-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="de518-105">場合によっては、Office アドインが、対応する COM アドインで使用可能なすべての機能を提供できないことがあります。</span><span class="sxs-lookup"><span data-stu-id="de518-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="de518-106">このような状況では、対応する Office アドインが提供するよりも、COM アドインによって Windows のユーザーの利便性が向上することがあります。</span><span class="sxs-lookup"><span data-stu-id="de518-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="de518-107">同等の COM アドインがユーザーのコンピューターに既にインストールされている場合に office アドインを構成すると、office アドインではなく、Windows が COM アドインを実行するようになります。</span><span class="sxs-lookup"><span data-stu-id="de518-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="de518-108">COM アドインは、Office がユーザーのコンピューターにインストールされているものに応じて、COM アドインと Office アドインをシームレスに移行するため、"同等" と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="de518-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="de518-109">この機能は、Office 365 サブスクリプションに接続されている場合、次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="de518-109">This feature is supported by the following platforms, when connected to an Office 365 subscription:</span></span>
> - <span data-ttu-id="de518-110">Excel、Word、および PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="de518-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="de518-111">Excel、Word、および PowerPoint on Windows (バージョン1904以降)</span><span class="sxs-lookup"><span data-stu-id="de518-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="de518-112">Excel、Word、および PowerPoint on Mac (バージョン13.329 以降)</span><span class="sxs-lookup"><span data-stu-id="de518-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="de518-113">マニフェストで同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="de518-113">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="de518-114">Office アドインと COM アドインの互換性を有効にするには、Office アドインの[マニフェスト](add-in-manifests.md)で同等の COM アドインを特定します。</span><span class="sxs-lookup"><span data-stu-id="de518-114">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="de518-115">その後、office アドインの両方がインストールされている場合は、Windows で office アドインではなく COM アドインが使用されます。</span><span class="sxs-lookup"><span data-stu-id="de518-115">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="de518-116">次の例は、COM アドインを同等のアドインとして指定するマニフェストの一部を示しています。</span><span class="sxs-lookup"><span data-stu-id="de518-116">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="de518-117">`ProgId`要素の値は COM アドインを識別し、 `EquivalentAddins`要素は終了`VersionOverrides`タグの直前に配置する必要があります。</span><span class="sxs-lookup"><span data-stu-id="de518-117">The value of the `ProgId` element identifies the COM add-in and the `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="de518-118">COM アドインと XLL UDF の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="de518-118">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="de518-119">ユーザーの同等の動作</span><span class="sxs-lookup"><span data-stu-id="de518-119">Equivalent behavior for users</span></span>

<span data-ttu-id="de518-120">Office アドインマニフェストで同等の COM アドインが指定されている場合、Windows 上の Office では、対応する COM アドインがインストールされている場合、Office アドインのユーザーインターフェイス (UI) は表示されません。</span><span class="sxs-lookup"><span data-stu-id="de518-120">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="de518-121">Office は、Office アドインのリボンボタンを非表示にし、インストールを妨げることはありません。</span><span class="sxs-lookup"><span data-stu-id="de518-121">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="de518-122">そのため、Office アドインは引き続き UI 内の次の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="de518-122">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="de518-123">[**個人用アドイン] の**下</span><span class="sxs-lookup"><span data-stu-id="de518-123">Under **My add-ins**</span></span>
- <span data-ttu-id="de518-124">リボンマネージャーのエントリとして</span><span class="sxs-lookup"><span data-stu-id="de518-124">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="de518-125">マニフェストで同等の COM アドインを指定しても、web または Mac の Office などの他のプラットフォームには影響しません。</span><span class="sxs-lookup"><span data-stu-id="de518-125">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Mac.</span></span>

<span data-ttu-id="de518-126">次のシナリオでは、ユーザーが Office アドインを取得する方法によって実行される処理について説明します。</span><span class="sxs-lookup"><span data-stu-id="de518-126">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="de518-127">Office アドインの AppSource 取得</span><span class="sxs-lookup"><span data-stu-id="de518-127">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="de518-128">ユーザーが AppSource から Office アドインを取得し、対応する COM アドインが既にインストールされている場合、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="de518-128">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="de518-129">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="de518-129">Install the Office Add-in.</span></span>
2. <span data-ttu-id="de518-130">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="de518-130">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="de518-131">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="de518-131">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="de518-132">Office アドインの一元展開</span><span class="sxs-lookup"><span data-stu-id="de518-132">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="de518-133">管理者が一元展開を使用して Office アドインをテナントに展開しており、対応する COM アドインが既にインストールされている場合、ユーザーは Office を再起動して変更を表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="de518-133">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="de518-134">Office を再起動すると、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="de518-134">After Office restarts, it will:</span></span>

1. <span data-ttu-id="de518-135">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="de518-135">Install the Office Add-in.</span></span>
2. <span data-ttu-id="de518-136">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="de518-136">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="de518-137">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="de518-137">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="de518-138">埋め込まれた Office アドインと共有されたドキュメント</span><span class="sxs-lookup"><span data-stu-id="de518-138">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="de518-139">ユーザーが COM アドインをインストールしていて、Office アドインが埋め込まれた共有ドキュメントを取得した場合、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="de518-139">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="de518-140">Office アドインを信頼するかどうかをユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="de518-140">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="de518-141">信頼できる場合は、Office アドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="de518-141">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="de518-142">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="de518-142">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="de518-143">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="de518-143">Other COM add-in behavior</span></span>

<span data-ttu-id="de518-144">ユーザーが同等の COM アドインをアンインストールした場合は、Windows の Office によって Office アドインの UI が復元されます。</span><span class="sxs-lookup"><span data-stu-id="de518-144">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="de518-145">Office アドインに対応する COM アドインを指定した後、office アドインの更新プログラムの処理を停止します。</span><span class="sxs-lookup"><span data-stu-id="de518-145">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="de518-146">Office アドインの最新の更新プログラムを入手するには、まず COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="de518-146">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="de518-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="de518-147">See also</span></span>

- [<span data-ttu-id="de518-148">カスタム関数を XLL ユーザー定義関数と互換性を持つようにする</span><span class="sxs-lookup"><span data-stu-id="de518-148">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
