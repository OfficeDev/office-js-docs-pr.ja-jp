---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bada3cd4cff7973517aedb83235a224ef6c273eb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901963"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="e9c2f-102">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="e9c2f-102">DesktopFormFactor element</span></span>

<span data-ttu-id="e9c2f-103">デスクトップフォームファクター用のアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="e9c2f-104">デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="e9c2f-105">このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="e9c2f-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="e9c2f-108">子要素</span><span class="sxs-lookup"><span data-stu-id="e9c2f-108">Child elements</span></span>

| <span data-ttu-id="e9c2f-109">要素</span><span class="sxs-lookup"><span data-stu-id="e9c2f-109">Element</span></span>                               | <span data-ttu-id="e9c2f-110">必須</span><span class="sxs-lookup"><span data-stu-id="e9c2f-110">Required</span></span> | <span data-ttu-id="e9c2f-111">説明</span><span class="sxs-lookup"><span data-stu-id="e9c2f-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="e9c2f-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="e9c2f-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="e9c2f-113">はい</span><span class="sxs-lookup"><span data-stu-id="e9c2f-113">Yes</span></span>      | <span data-ttu-id="e9c2f-114">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="e9c2f-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="e9c2f-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="e9c2f-116">はい</span><span class="sxs-lookup"><span data-stu-id="e9c2f-116">Yes</span></span>      | <span data-ttu-id="e9c2f-117">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="e9c2f-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="e9c2f-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="e9c2f-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9c2f-119">No</span></span>       | <span data-ttu-id="e9c2f-120">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="e9c2f-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e9c2f-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="e9c2f-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="e9c2f-122">No</span></span> | <span data-ttu-id="e9c2f-123">代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。</span><span class="sxs-lookup"><span data-stu-id="e9c2f-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="e9c2f-124">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="e9c2f-124">DesktopFormFactor example</span></span>

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
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
