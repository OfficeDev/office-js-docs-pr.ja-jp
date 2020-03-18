---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: デスクトップフォームファクター用のアドインの設定を指定します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bfea6900e6b07d8dc7ad5b5256703d873242d88c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718364"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="6f3ad-103">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="6f3ad-103">DesktopFormFactor element</span></span>

<span data-ttu-id="6f3ad-104">デスクトップフォームファクター用のアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="6f3ad-105">デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="6f3ad-106">このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="6f3ad-107">各 DesktopFormFactor 定義には、 **Functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="6f3ad-108">詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="6f3ad-109">子要素</span><span class="sxs-lookup"><span data-stu-id="6f3ad-109">Child elements</span></span>

| <span data-ttu-id="6f3ad-110">要素</span><span class="sxs-lookup"><span data-stu-id="6f3ad-110">Element</span></span>                               | <span data-ttu-id="6f3ad-111">必須</span><span class="sxs-lookup"><span data-stu-id="6f3ad-111">Required</span></span> | <span data-ttu-id="6f3ad-112">説明</span><span class="sxs-lookup"><span data-stu-id="6f3ad-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="6f3ad-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="6f3ad-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="6f3ad-114">はい</span><span class="sxs-lookup"><span data-stu-id="6f3ad-114">Yes</span></span>      | <span data-ttu-id="6f3ad-115">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="6f3ad-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="6f3ad-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="6f3ad-117">はい</span><span class="sxs-lookup"><span data-stu-id="6f3ad-117">Yes</span></span>      | <span data-ttu-id="6f3ad-118">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="6f3ad-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="6f3ad-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="6f3ad-120">いいえ</span><span class="sxs-lookup"><span data-stu-id="6f3ad-120">No</span></span>       | <span data-ttu-id="6f3ad-121">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="6f3ad-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="6f3ad-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="6f3ad-123">いいえ</span><span class="sxs-lookup"><span data-stu-id="6f3ad-123">No</span></span> | <span data-ttu-id="6f3ad-124">代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。</span><span class="sxs-lookup"><span data-stu-id="6f3ad-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="6f3ad-125">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="6f3ad-125">DesktopFormFactor example</span></span>

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
