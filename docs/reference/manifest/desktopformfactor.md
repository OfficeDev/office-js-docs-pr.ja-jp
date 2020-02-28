---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2fe97d99ff5bdc9f23a5760824e241ee4dfb800f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325277"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="44787-102">DesktopFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="44787-102">DesktopFormFactor element</span></span>

<span data-ttu-id="44787-103">デスクトップフォームファクター用のアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="44787-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="44787-104">デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。</span><span class="sxs-lookup"><span data-stu-id="44787-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="44787-105">このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。</span><span class="sxs-lookup"><span data-stu-id="44787-105">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="44787-106">各 DesktopFormFactor 定義には、 **Functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="44787-106">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="44787-107">詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44787-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="44787-108">子要素</span><span class="sxs-lookup"><span data-stu-id="44787-108">Child elements</span></span>

| <span data-ttu-id="44787-109">要素</span><span class="sxs-lookup"><span data-stu-id="44787-109">Element</span></span>                               | <span data-ttu-id="44787-110">必須</span><span class="sxs-lookup"><span data-stu-id="44787-110">Required</span></span> | <span data-ttu-id="44787-111">説明</span><span class="sxs-lookup"><span data-stu-id="44787-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="44787-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="44787-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="44787-113">はい</span><span class="sxs-lookup"><span data-stu-id="44787-113">Yes</span></span>      | <span data-ttu-id="44787-114">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="44787-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="44787-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="44787-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="44787-116">はい</span><span class="sxs-lookup"><span data-stu-id="44787-116">Yes</span></span>      | <span data-ttu-id="44787-117">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="44787-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="44787-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="44787-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="44787-119">いいえ</span><span class="sxs-lookup"><span data-stu-id="44787-119">No</span></span>       | <span data-ttu-id="44787-120">Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。</span><span class="sxs-lookup"><span data-stu-id="44787-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="44787-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="44787-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="44787-122">いいえ</span><span class="sxs-lookup"><span data-stu-id="44787-122">No</span></span> | <span data-ttu-id="44787-123">代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。</span><span class="sxs-lookup"><span data-stu-id="44787-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="44787-124">DesktopFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="44787-124">DesktopFormFactor example</span></span>

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
