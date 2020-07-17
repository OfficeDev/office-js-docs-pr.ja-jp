---
title: Office アドインを既存の COM アドインと互換できるようにする
description: Office アドインと同等の COM アドインの互換性を有効にする
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: 1ded60366c36edb876b8b2dd0a4ea9e6a39521b0
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093596"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="21780-103">Office アドインを既存の COM アドインと互換できるようにする</span><span class="sxs-lookup"><span data-stu-id="21780-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="21780-104">既存の COM アドインがある場合は、Office アドインで同等の機能を構築できます。これにより、web や Office on the Mac 上の他のプラットフォーム上でソリューションを実行することが可能になります。</span><span class="sxs-lookup"><span data-stu-id="21780-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in, thereby enabling your solution to run on other platforms such as Office on the web or Office on Mac.</span></span> <span data-ttu-id="21780-105">場合によっては、Office アドインが、対応する COM アドインで使用可能なすべての機能を提供できないことがあります。</span><span class="sxs-lookup"><span data-stu-id="21780-105">In some cases, your Office Add-in may not be able to provide all of the functionality that's available in the corresponding COM add-in.</span></span> <span data-ttu-id="21780-106">このような状況では、対応する Office アドインが提供するよりも、COM アドインによって Windows のユーザーの利便性が向上することがあります。</span><span class="sxs-lookup"><span data-stu-id="21780-106">In these situations, your COM add-in may provide a better user experience on Windows than the corresponding Office Add-in can provide.</span></span>

<span data-ttu-id="21780-107">同等の COM アドインがユーザーのコンピューターに既にインストールされている場合に office アドインを構成すると、office アドインではなく、Windows が COM アドインを実行するようになります。</span><span class="sxs-lookup"><span data-stu-id="21780-107">You can configure your Office Add-in so that when the equivalent COM add-in is already installed on a user's computer, Office on Windows runs the COM add-in instead of the Office Add-in.</span></span> <span data-ttu-id="21780-108">COM アドインは、Office がユーザーのコンピューターにインストールされているものに応じて、COM アドインと Office アドインをシームレスに移行するため、"同等" と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="21780-108">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in according to which one is installed a user's computer.</span></span>

> [!NOTE]
> <span data-ttu-id="21780-109">この機能は、Microsoft 365 サブスクリプションに接続する際に、次のプラットフォームでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="21780-109">This feature is supported by the following platforms, when connected to a Microsoft 365 subscription:</span></span>
> - <span data-ttu-id="21780-110">Excel、Word、および PowerPoint on the web</span><span class="sxs-lookup"><span data-stu-id="21780-110">Excel, Word, and PowerPoint on the web</span></span>
> - <span data-ttu-id="21780-111">Excel、Word、および PowerPoint on Windows (バージョン1904以降)</span><span class="sxs-lookup"><span data-stu-id="21780-111">Excel, Word, and PowerPoint on Windows (version 1904 or later)</span></span>
> - <span data-ttu-id="21780-112">Excel、Word、および PowerPoint on Mac (バージョン13.329 以降)</span><span class="sxs-lookup"><span data-stu-id="21780-112">Excel, Word, and PowerPoint on Mac (version 13.329 or later)</span></span>

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="21780-113">マニフェストで同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="21780-113">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="21780-114">Office アドインと COM アドインの互換性を有効にするには、Office アドインの[マニフェスト](add-in-manifests.md)で同等の COM アドインを特定します。</span><span class="sxs-lookup"><span data-stu-id="21780-114">To enable compatibility between your Office Add-in and COM add-in, identify the equivalent COM add-in in the [manifest](add-in-manifests.md) of your Office Add-in.</span></span> <span data-ttu-id="21780-115">その後、office アドインの両方がインストールされている場合は、Windows で office アドインではなく COM アドインが使用されます。</span><span class="sxs-lookup"><span data-stu-id="21780-115">Then Office on Windows will use the COM add-in instead of the Office Add-in, if they're both installed.</span></span>

<span data-ttu-id="21780-116">次の例は、COM アドインを同等のアドインとして指定するマニフェストの一部を示しています。</span><span class="sxs-lookup"><span data-stu-id="21780-116">The following example shows the portion of the manifest that specifies a COM add-in as an equivalent add-in.</span></span> <span data-ttu-id="21780-117">要素の値は `ProgId` COM アドインを識別し、 `EquivalentAddins` 要素は終了タグの直前に配置する必要があり `VersionOverrides` ます。</span><span class="sxs-lookup"><span data-stu-id="21780-117">The value of the `ProgId` element identifies the COM add-in and the `EquivalentAddins` element must be positioned immediately before the closing `VersionOverrides` tag.</span></span>

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
> <span data-ttu-id="21780-118">COM アドインと XLL UDF の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="21780-118">For information about COM add-in and XLL UDF compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="21780-119">ユーザーの同等の動作</span><span class="sxs-lookup"><span data-stu-id="21780-119">Equivalent behavior for users</span></span>

<span data-ttu-id="21780-120">Office アドインマニフェストで同等の COM アドインが指定されている場合、Windows 上の Office では、対応する COM アドインがインストールされている場合、Office アドインのユーザーインターフェイス (UI) は表示されません。</span><span class="sxs-lookup"><span data-stu-id="21780-120">When an equivalent COM add-in is specified in the Office Add-in manifest, Office on Windows will not display your Office Add-in's user interface (UI) if the equivalent COM add-in is installed.</span></span> <span data-ttu-id="21780-121">Office は、Office アドインのリボンボタンを非表示にし、インストールを妨げることはありません。</span><span class="sxs-lookup"><span data-stu-id="21780-121">Office only hides the ribbon buttons of the Office Add-in and does not prevent installation.</span></span> <span data-ttu-id="21780-122">そのため、Office アドインは引き続き UI 内の次の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="21780-122">Therefore your Office Add-in will still appear in the following locations within the UI:</span></span>

- <span data-ttu-id="21780-123">[**個人用アドイン] の**下</span><span class="sxs-lookup"><span data-stu-id="21780-123">Under **My add-ins**</span></span>
- <span data-ttu-id="21780-124">リボンマネージャーのエントリとして</span><span class="sxs-lookup"><span data-stu-id="21780-124">As an entry in the ribbon manager</span></span>

> [!NOTE]
> <span data-ttu-id="21780-125">マニフェストで同等の COM アドインを指定しても、web または Mac の Office などの他のプラットフォームには影響しません。</span><span class="sxs-lookup"><span data-stu-id="21780-125">Specifying an equivalent COM add-in in the manifest has no effect on other platforms like Office on the web or Mac.</span></span>

<span data-ttu-id="21780-126">次のシナリオでは、ユーザーが Office アドインを取得する方法によって実行される処理について説明します。</span><span class="sxs-lookup"><span data-stu-id="21780-126">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="21780-127">Office アドインの AppSource 取得</span><span class="sxs-lookup"><span data-stu-id="21780-127">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="21780-128">ユーザーが AppSource から Office アドインを取得し、対応する COM アドインが既にインストールされている場合、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="21780-128">If a user acquires the Office Add-in from AppSource and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="21780-129">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="21780-129">Install the Office Add-in.</span></span>
2. <span data-ttu-id="21780-130">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="21780-130">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="21780-131">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="21780-131">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="21780-132">Office アドインの一元展開</span><span class="sxs-lookup"><span data-stu-id="21780-132">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="21780-133">管理者が一元展開を使用して Office アドインをテナントに展開しており、対応する COM アドインが既にインストールされている場合、ユーザーは Office を再起動して変更を表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="21780-133">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user must restart Office before they'll see any changes.</span></span> <span data-ttu-id="21780-134">Office を再起動すると、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="21780-134">After Office restarts, it will:</span></span>

1. <span data-ttu-id="21780-135">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="21780-135">Install the Office Add-in.</span></span>
2. <span data-ttu-id="21780-136">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="21780-136">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="21780-137">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="21780-137">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="21780-138">埋め込まれた Office アドインと共有されたドキュメント</span><span class="sxs-lookup"><span data-stu-id="21780-138">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="21780-139">ユーザーが COM アドインをインストールしていて、Office アドインが埋め込まれた共有ドキュメントを取得した場合、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="21780-139">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="21780-140">Office アドインを信頼するかどうかをユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="21780-140">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="21780-141">信頼できる場合は、Office アドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="21780-141">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="21780-142">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="21780-142">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="21780-143">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="21780-143">Other COM add-in behavior</span></span>

<span data-ttu-id="21780-144">ユーザーが同等の COM アドインをアンインストールした場合は、Windows の Office によって Office アドインの UI が復元されます。</span><span class="sxs-lookup"><span data-stu-id="21780-144">If a user uninstalls the equivalent COM add-in, then Office on Windows restores the Office Add-in UI.</span></span>

<span data-ttu-id="21780-145">Office アドインに対応する COM アドインを指定した後、office アドインの更新プログラムの処理を停止します。</span><span class="sxs-lookup"><span data-stu-id="21780-145">After you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="21780-146">Office アドインの最新の更新プログラムを入手するには、まず COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="21780-146">To acquire the latest updates for the Office Add-in, the user must first uninstall the COM add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="21780-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="21780-147">See also</span></span>

- [<span data-ttu-id="21780-148">カスタム関数を XLL ユーザー定義関数と互換性を持つようにする</span><span class="sxs-lookup"><span data-stu-id="21780-148">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
