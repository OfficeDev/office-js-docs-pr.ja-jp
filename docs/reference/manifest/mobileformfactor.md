---
title: マニフェスト ファイルの MobileFormFactor 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324842"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="fdaa8-102">MobileFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="fdaa8-102">MobileFormFactor element</span></span>

<span data-ttu-id="fdaa8-p101">モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="fdaa8-105">各**MobileFormFactor**定義には、 **functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-105">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="fdaa8-106">詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-106">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="fdaa8-p103">**MobileFormFactor** 要素は、VersionOverrides のスキーマ 1.1 で定義されています。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="fdaa8-109">子要素</span><span class="sxs-lookup"><span data-stu-id="fdaa8-109">Child elements</span></span>

| <span data-ttu-id="fdaa8-110">要素</span><span class="sxs-lookup"><span data-stu-id="fdaa8-110">Element</span></span>                               | <span data-ttu-id="fdaa8-111">必須</span><span class="sxs-lookup"><span data-stu-id="fdaa8-111">Required</span></span> | <span data-ttu-id="fdaa8-112">説明</span><span class="sxs-lookup"><span data-stu-id="fdaa8-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="fdaa8-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="fdaa8-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="fdaa8-114">はい</span><span class="sxs-lookup"><span data-stu-id="fdaa8-114">Yes</span></span>      | <span data-ttu-id="fdaa8-115">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="fdaa8-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="fdaa8-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="fdaa8-117">はい</span><span class="sxs-lookup"><span data-stu-id="fdaa8-117">Yes</span></span>      | <span data-ttu-id="fdaa8-118">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="fdaa8-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="fdaa8-119">MobileFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="fdaa8-119">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
