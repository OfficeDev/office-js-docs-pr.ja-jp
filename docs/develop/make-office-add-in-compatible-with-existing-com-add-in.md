---
title: 既存の COM アドインと互換性のある Excel アドインを作成する
description: Excel アドインと同じ機能を持つ同等の COM アドインとの互換性を有効にする
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628173"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a><span data-ttu-id="4c31e-103">既存の COM アドインと互換性のある Office アドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="4c31e-103">Make your Office Add-in compatible with an existing COM add-in (preview)</span></span>

<span data-ttu-id="4c31e-104">既存の COM アドインがある場合は、Excel アドインで同等の機能を構築して、ソリューション機能をオンラインや macOS などの他のプラットフォームに拡張できます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-104">If you have an existing COM add-in, you can build equivalent functionality in your Excel add-in to extend your solution features to other platforms such as online or macOS.</span></span> <span data-ttu-id="4c31e-105">ただし、Excel アドインには、COM アドインで使用できるすべての機能が含まれているわけではありません。COM アドインを使用すると、Windows の Excel アドインよりも優れたパフォーマンスを得ることができます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-105">However, Excel add-ins don't have all of the functionality available in COM add-ins. Your COM add-in may provide a better experience than the Excel add-in on Windows.</span></span>

<span data-ttu-id="4c31e-106">同等の COM アドインがユーザーのコンピューターに既にインストールされている場合、Office は Excel アドインではなく COM アドインを実行するように、Excel アドインを構成することができます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-106">You can configure your Excel add-in so that when an equivalent COM add-in is already installed on the user's computer, Office runs the COM add-in instead of your Excel add-in.</span></span> <span data-ttu-id="4c31e-107">COM アドインは、Windows にインストールされているものに応じて、COM アドインと Excel アドインの間でシームレスに移行されるため、"同等" と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-107">The COM add-in is called "equivalent" because Office will seamlessly transition between the COM add-in and the Excel add-in depending on which is installed on Windows.</span></span>

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a><span data-ttu-id="4c31e-108">マニフェストで同等の COM アドインを指定する</span><span class="sxs-lookup"><span data-stu-id="4c31e-108">Specify an equivalent COM add-in in the manifest</span></span>

<span data-ttu-id="4c31e-109">既存の COM アドインとの互換性を有効にするには、Excel アドインのマニフェストで同等の COM アドインを特定します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-109">To enable compatibility with an existing COM add-in, identify the equivalent COM add-in in the manifest of your Excel add-in.</span></span> <span data-ttu-id="4c31e-110">Windows で実行している場合、Office は Excel アドインではなく COM アドインを使用します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-110">Then Office will use the COM add-in instead of your Excel add-in when running on Windows.</span></span>

<span data-ttu-id="4c31e-111">同等の`ProgID` COM アドインのを指定します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-111">Specify the `ProgID` of the equivalent COM add-in.</span></span> <span data-ttu-id="4c31e-112">COM アドインがインストールされている場合、Office は、Excel アドインの UI ではなく、COM アドインの UI を使用します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-112">Office will then use the COM add-in UI instead of your Excel add-in's UI when the COM add-in is installed.</span></span>

<span data-ttu-id="4c31e-113">次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="4c31e-113">The following example shows how to specify both a COM add-in and an XLL as equivalent.</span></span> <span data-ttu-id="4c31e-114">多くの場合、この例は完全にコンテキストで指定します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-114">Often you will specify both so for completeness this example shows both in context.</span></span> <span data-ttu-id="4c31e-115">これらは、 `FileName`それぞれに`ProgID`よって識別されます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-115">They are identified by their `ProgID` and `FileName` respectively.</span></span> <span data-ttu-id="4c31e-116">XLL の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4c31e-116">For more information on XLL compatibility, see [Make your custom functions compatible with XLL user-defined functions](../excel/make-custom-functions-compatible-with-xll-udf.md).</span></span>

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

## <a name="equivalent-behavior-for-users"></a><span data-ttu-id="4c31e-117">ユーザーの同等の動作</span><span class="sxs-lookup"><span data-stu-id="4c31e-117">Equivalent behavior for users</span></span>

<span data-ttu-id="4c31e-118">同等の COM アドインが Excel アドインマニフェストで指定されている場合、Office は同等の COM アドインがインストールされている場合、Windows 上で Excel アドインの UI を非表示にします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-118">When an equivalent COM add-in is specified in the Excel add-in manifest, Office suppresses your Excel add-in's UI on Windows when the equivalent COM add-in is installed.</span></span> <span data-ttu-id="4c31e-119">これは、オンラインまたは macOS などの他のプラットフォームで Excel アドインの UI に影響を与えることはありません。</span><span class="sxs-lookup"><span data-stu-id="4c31e-119">This does not affect your Excel add-in's UI on other platforms like online or macOS.</span></span> <span data-ttu-id="4c31e-120">Office はリボンボタンを非表示にし、インストールを妨げることはありません。</span><span class="sxs-lookup"><span data-stu-id="4c31e-120">Office only hides the ribbon buttons and does not prevent installation.</span></span> <span data-ttu-id="4c31e-121">そのため、Excel アドインは引き続き次の UI の場所に表示されます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-121">Therefore your Excel add-in will still appear in the following UI locations:</span></span>

- <span data-ttu-id="4c31e-122">[ \*\*\*\* アドイン] の下で、技術的にインストールされています。</span><span class="sxs-lookup"><span data-stu-id="4c31e-122">Under **My add-ins** because it is technically installed.</span></span>
- <span data-ttu-id="4c31e-123">リボンマネージャーのエントリとして。</span><span class="sxs-lookup"><span data-stu-id="4c31e-123">As an entry in the ribbon manager.</span></span>

<span data-ttu-id="4c31e-124">次のシナリオでは、ユーザーが Excel アドインを取得する方法によって実行される処理について説明します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-124">The following scenarios describe what happens depending on how the user acquires the Excel add-in.</span></span>

### <a name="appsource-acquisition-of-an-excel-add-in"></a><span data-ttu-id="4c31e-125">Excel アドインの AppSource 取得</span><span class="sxs-lookup"><span data-stu-id="4c31e-125">AppSource acquisition of an Excel add-in</span></span>

<span data-ttu-id="4c31e-126">ユーザーが AppSource から Excel アドインをダウンロードし、対応する COM アドインが既にインストールされている場合、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="4c31e-126">If a user downloads the Excel add-in from AppSource, and the equivalent COM add-in is already installed, then Office will:</span></span>

1. <span data-ttu-id="4c31e-127">Excel アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-127">Install the Excel add-in.</span></span>
2. <span data-ttu-id="4c31e-128">リボンに Excel アドイン UI を表示しないようにします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-128">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="4c31e-129">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-129">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="centralized-deployment-of-excel-add-in"></a><span data-ttu-id="4c31e-130">Excel アドインの一元展開</span><span class="sxs-lookup"><span data-stu-id="4c31e-130">Centralized deployment of Excel add-in</span></span>

<span data-ttu-id="4c31e-131">管理者が一元展開を使用して Excel アドインをテナントに展開していて、それと同等の COM アドインが既にインストールされている場合、ユーザーは変更を確認する前に Office を再起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c31e-131">If an admin deploys the Excel add-in to their tenant using centralized deployment, and the equivalent COM add-in is already installed, the user needs to restart Office before they will see any changes.</span></span> <span data-ttu-id="4c31e-132">Office を再起動すると、次のようになります。</span><span class="sxs-lookup"><span data-stu-id="4c31e-132">After Office restarts, it will:</span></span>

1. <span data-ttu-id="4c31e-133">Excel アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-133">Install the Excel add-in.</span></span>
2. <span data-ttu-id="4c31e-134">リボンに Excel アドイン UI を表示しないようにします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-134">Hide the Excel add-in UI in the ribbon.</span></span>
3. <span data-ttu-id="4c31e-135">[COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-135">Display a call-out for the user that points out the COM add-in ribbon button.</span></span>

### <a name="document-shared-with-embedded-excel-add-in"></a><span data-ttu-id="4c31e-136">埋め込まれた Excel アドインで共有されたドキュメント</span><span class="sxs-lookup"><span data-stu-id="4c31e-136">Document shared with embedded Excel add-in</span></span>

<span data-ttu-id="4c31e-137">ユーザーが COM アドインをインストールしている場合に、埋め込まれた Excel アドインを含む共有ドキュメントを取得すると、Office は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="4c31e-137">If a user has the COM add-in installed, and then gets a shared document with the embedded Excel add-in, then when they open the document, Office will:</span></span>

1. <span data-ttu-id="4c31e-138">Excel アドインを信頼するかどうかをユーザーに確認します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-138">Prompt the user to trust the Excel add-in.</span></span>
2. <span data-ttu-id="4c31e-139">信頼できる場合は、Excel アドインがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="4c31e-139">If trusted, the Excel add-in will install.</span></span>
3. <span data-ttu-id="4c31e-140">リボンに Excel アドイン UI を表示しないようにします。</span><span class="sxs-lookup"><span data-stu-id="4c31e-140">Hide the Excel add-in UI in the ribbon.</span></span>

## <a name="other-com-add-in-behavior"></a><span data-ttu-id="4c31e-141">その他の COM アドインの動作</span><span class="sxs-lookup"><span data-stu-id="4c31e-141">Other COM add-in behavior</span></span>

<span data-ttu-id="4c31e-142">ユーザーが COM アドインをアンインストールすると、Office は、インストールされている excel アドインに対応する Excel アドインの UI を Windows 上に復元します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-142">If a user uninstalls the COM add-in, then Office restores the Excel add-in UI on Windows for the equivalent installed Excel add-in.</span></span>

<span data-ttu-id="4c31e-143">Excel アドインに対して同等の COM アドインを指定すると、Office は Excel アドインの更新プログラムの処理を停止します。</span><span class="sxs-lookup"><span data-stu-id="4c31e-143">Once you specify an equivalent COM add-in for your Excel add-in, Office stops processing updates for your Excel add-in.</span></span> <span data-ttu-id="4c31e-144">ユーザーは、Excel アドインの最新の更新プログラムを取得するために、COM アドインをアンインストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="4c31e-144">The user must uninstall the COM add-in order to get the latest updates for the Excel add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="4c31e-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="4c31e-145">See also</span></span>

- [<span data-ttu-id="4c31e-146">カスタム関数を XLL ユーザー定義関数と互換性を持つようにする</span><span class="sxs-lookup"><span data-stu-id="4c31e-146">Make your Custom Functions compatible with XLL User Defined Functions</span></span>](../excel/make-custom-functions-compatible-with-xll-udf.md)
