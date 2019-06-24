---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d1f09203518a38f1568b13e6c1a9c70752697152
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128518"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="7bd62-102">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="7bd62-102">DesktopFormFactor element</span></span>

<span data-ttu-id="7bd62-103">デスクトップフォームファクター用のアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="7bd62-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="7bd62-104">デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7bd62-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="7bd62-105">このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7bd62-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="7bd62-p102">各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7bd62-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="7bd62-108">子要素</span><span class="sxs-lookup"><span data-stu-id="7bd62-108">Child elements</span></span>

| <span data-ttu-id="7bd62-109">要素</span><span class="sxs-lookup"><span data-stu-id="7bd62-109">Element</span></span>                               | <span data-ttu-id="7bd62-110">必須</span><span class="sxs-lookup"><span data-stu-id="7bd62-110">Required</span></span> | <span data-ttu-id="7bd62-111">説明</span><span class="sxs-lookup"><span data-stu-id="7bd62-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="7bd62-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="7bd62-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="7bd62-113">はい</span><span class="sxs-lookup"><span data-stu-id="7bd62-113">Yes</span></span>      | <span data-ttu-id="7bd62-114">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="7bd62-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="7bd62-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="7bd62-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="7bd62-116">はい</span><span class="sxs-lookup"><span data-stu-id="7bd62-116">Yes</span></span>      | <span data-ttu-id="7bd62-117">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="7bd62-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="7bd62-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="7bd62-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="7bd62-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="7bd62-119">No</span></span>       | <span data-ttu-id="7bd62-120">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="7bd62-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="7bd62-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="7bd62-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="7bd62-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="7bd62-122">No</span></span> | <span data-ttu-id="7bd62-123">代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。</span><span class="sxs-lookup"><span data-stu-id="7bd62-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="7bd62-124">**重要**: Outlook アドインの代理人アクセスは現在プレビュー段階であるため、この`SupportSharedFolders`要素を使用するアドインは、appsource に発行することも、一元展開によって展開することもできません。</span><span class="sxs-lookup"><span data-stu-id="7bd62-124">**Important**: Because delegate access for Outlook add-ins is currently in preview, add-ins that use the `SupportSharedFolders` element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="7bd62-125">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="7bd62-125">DesktopFormFactor example</span></span>

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
