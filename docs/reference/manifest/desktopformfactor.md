---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433741"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="280dd-102">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="280dd-102">DesktopFormFactor element</span></span>

<span data-ttu-id="280dd-p101">デスクトップ フォーム ファクターについてアドインの設定を指定します。デスクトップ フォーム ファクターには、Office for Windows、Office for Mac、Office Online が含まれています。**Resources** ノードを除くデスクトップ フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="280dd-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="280dd-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="280dd-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="280dd-108">子要素</span><span class="sxs-lookup"><span data-stu-id="280dd-108">Child elements</span></span>

| <span data-ttu-id="280dd-109">要素</span><span class="sxs-lookup"><span data-stu-id="280dd-109">Element</span></span>                               | <span data-ttu-id="280dd-110">必須</span><span class="sxs-lookup"><span data-stu-id="280dd-110">Required</span></span> | <span data-ttu-id="280dd-111">説明</span><span class="sxs-lookup"><span data-stu-id="280dd-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="280dd-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="280dd-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="280dd-113">はい</span><span class="sxs-lookup"><span data-stu-id="280dd-113">Yes</span></span>      | <span data-ttu-id="280dd-114">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="280dd-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="280dd-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="280dd-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="280dd-116">はい</span><span class="sxs-lookup"><span data-stu-id="280dd-116">Yes</span></span>      | <span data-ttu-id="280dd-117">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="280dd-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="280dd-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="280dd-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="280dd-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="280dd-119">No</span></span>       | <span data-ttu-id="280dd-120">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="280dd-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="280dd-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="280dd-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="280dd-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="280dd-122">No</span></span> | <span data-ttu-id="280dd-123">代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。</span><span class="sxs-lookup"><span data-stu-id="280dd-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="280dd-124">**重要事項**: この要素は、Outlook アドイン要件セットのプレビューの Exchange Online に対してのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="280dd-124">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="280dd-125">この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="280dd-125">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="280dd-126">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="280dd-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
