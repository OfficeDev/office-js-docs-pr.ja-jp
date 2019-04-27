---
title: 既存の COM アドインと互換性のある Office アドインを作成する
description: Office アドインと同じ機能を持つ同等の COM アドインとの互換性を有効にする
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356895"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a><span data-ttu-id="674ec-103">既存の COM アドインと互換性のある Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="674ec-103">Make your Office Add-in compatible with an existing COM add-in</span></span>

<span data-ttu-id="674ec-104">既存の COM アドインがある場合は、Office アドインで同等の機能を構築して、ソリューション機能を online や macOS などの他のプラットフォームに拡張できます。</span><span class="sxs-lookup"><span data-stu-id="674ec-104">If you have an existing COM add-in, you can build equivalent functionality in your Office Add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="674ec-105">ただし、Office アドインには、COM アドインで使用できるすべての機能が含まれているわけではありません。COM アドインでは、Excel、Word、および PowerPoint の Office アドインよりも優れた機能を提供する場合があります。</span><span class="sxs-lookup"><span data-stu-id="674ec-105">However, Office Add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Office Add-in on Windows in Excel, Word, and PowerPoint.</span></span>

<span data-ttu-id="674ec-106">同等の com アドインがユーザーのコンピューターに既にインストールされている場合は office アドインを構成できます。 office は、office アドインではなく、com アドインを実行します。</span><span class="sxs-lookup"><span data-stu-id="674ec-106">You can configure your Office Add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Office Add-in.</span></span> <span data-ttu-id="674ec-107">com アドインは、office が Windows にインストールされているものに応じて、com アドインと office アドインをシームレスに移行するため、"同等" と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="674ec-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Office Add-in depending on which is installed on Windows.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="674ec-108">マニフェストで同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="674ec-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="674ec-109">既存の com アドインとの互換性を有効にするには、Office アドインのマニフェストで同等の com アドインを特定します。</span><span class="sxs-lookup"><span data-stu-id="674ec-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Office Add-in.</span></span> <span data-ttu-id="674ec-110">Windows で実行している場合、office は office アドインではなく、COM アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="674ec-110">Then Office will use the COM add-in instead of your Office Add-in when running on Windows.</span></span>

<span data-ttu-id="674ec-111">同等の`ProgID` COM アドインのを指定します。</span><span class="sxs-lookup"><span data-stu-id="674ec-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="674ec-112">これで、com アドインをインストールするときに、office アドインの ui ではなく、com アドインの ui が使用されます。</span><span class="sxs-lookup"><span data-stu-id="674ec-112">Office will then use the COM add-in UI instead of your Office Add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="674ec-113">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="674ec-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="674ec-114">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="674ec-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="674ec-115">これらは、 `FileName`それぞれに`ProgID`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="674ec-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="674ec-116">xll の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="674ec-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="674ec-117">ユーザーの同等の動作</span><span class="sxs-lookup"><span data-stu-id="674ec-117">Equivalent behavior for users</span></span>

<span data-ttu-id="674ec-118">office アドインマニフェストで同等の com アドインが指定されている場合、office は、対応する com アドインがインストールされている場合、Windows 上で office アドインの UI を表示しません。</span><span class="sxs-lookup"><span data-stu-id="674ec-118">When an equivalent COM add-in is specified in the Office Add-in manifest, Office suppresses your Office Add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="674ec-119">これは、online や macOS などの他のプラットフォームで Office アドインの UI に影響を与えることはありません。</span><span class="sxs-lookup"><span data-stu-id="674ec-119">This does not affect your Office Add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="674ec-120">Office はリボンボタンを非表示にし、インストールを妨げることはありません。</span><span class="sxs-lookup"><span data-stu-id="674ec-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="674ec-121">そのため、Office アドインは引き続き次の UI の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="674ec-121">Therefore your Office Add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="674ec-122">[ \*\*\*\* アドイン] の下で、技術的にインストールされています。</span><span class="sxs-lookup"><span data-stu-id="674ec-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="674ec-123">リボンマネージャーのエントリとして。</span><span class="sxs-lookup"><span data-stu-id="674ec-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="674ec-124">次のシナリオでは、ユーザーが Office アドインを取得する方法によって実行される処理について説明します。</span><span class="sxs-lookup"><span data-stu-id="674ec-124">The following scenarios describe what happens depending on how the user acquires the Office Add-in.</span></span>

### <a name="appsource-acquisition-of-an-office-add-in"></a><span data-ttu-id="674ec-125">Office アドインの appsource 取得</span><span class="sxs-lookup"><span data-stu-id="674ec-125">AppSource acquisition of an Office Add-in</span></span>

<span data-ttu-id="674ec-126">ユーザーが appsource から Office アドインをダウンロードし、対応する COM アドインが既にインストールされている場合、office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="674ec-126">If a user downloads the Office Add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="674ec-127">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="674ec-127">Install the Office Add-in.</span></span>
2. <span data-ttu-id="674ec-128">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="674ec-128">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="674ec-129">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="674ec-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-office-add-in"></a><span data-ttu-id="674ec-130">Office アドインの一元展開</span><span class="sxs-lookup"><span data-stu-id="674ec-130">Centralized deployment of Office Add-in</span></span>

<span data-ttu-id="674ec-131">管理者が一元展開を使用して office アドインをテナントに展開していて、それと同等の COM アドインが既にインストールされている場合、ユーザーは変更を確認する前に office を再起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="674ec-131">If an admin deploys the Office Add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="674ec-132">Office を再起動すると、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="674ec-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="674ec-133">Office アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="674ec-133">Install the Office Add-in.</span></span>
2. <span data-ttu-id="674ec-134">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="674ec-134">Hide the Office Add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="674ec-135">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="674ec-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-office-add-in"></a><span data-ttu-id="674ec-136">埋め込まれた Office アドインと共有されたドキュメント</span><span class="sxs-lookup"><span data-stu-id="674ec-136">Document shared with embedded Office Add-in</span></span>

<span data-ttu-id="674ec-137">ユーザーが COM アドインをインストールしていて、office アドインが埋め込まれた共有ドキュメントを取得した場合、office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="674ec-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Office Add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="674ec-138">Office アドインを信頼するかどうかをユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="674ec-138">Prompt the user to trust the Office Add-in.</span></span>
2. <span data-ttu-id="674ec-139">信頼できる場合は、Office アドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="674ec-139">If trusted, the Office Add-in will install.</span></span>
3. <span data-ttu-id="674ec-140">リボンで Office アドイン UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="674ec-140">Hide the Office Add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="674ec-141">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="674ec-141">Other COM add-in behavior</span></span>

<span data-ttu-id="674ec-142">ユーザーが COM アドインをアンインストールすると、office アドインの UI は、インストールされている office アドインに対応する Windows 上で復元されます。</span><span class="sxs-lookup"><span data-stu-id="674ec-142">If a user uninstalls the COM add-in, then Office restores the Office Add-in UI on Windows for the equivalent installed Office Add-in.</span></span>

<span data-ttu-id="674ec-143">office アドインに対して同等の COM アドインを指定すると、office アドインの更新プログラムの処理は中止されます。</span><span class="sxs-lookup"><span data-stu-id="674ec-143">Once you specify an equivalent COM add-in for your Office Add-in, Office stops processing updates for your Office Add-in.</span></span> <span data-ttu-id="674ec-144">ユーザーは、Office アドインの最新の更新プログラムを取得するために、COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="674ec-144">The user must uninstall the COM add-in order to get the latest updates for the Office Add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="674ec-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="674ec-145">See also</span></span>

- [<span data-ttu-id="674ec-146">カスタム関数を XLL ユーザー定義関数と互換性を持つようにする</span><span class="sxs-lookup"><span data-stu-id="674ec-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
