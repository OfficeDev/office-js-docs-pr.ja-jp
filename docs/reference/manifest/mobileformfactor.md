---
title: マニフェスト ファイルの MobileFormFactor 要素
description: MobileFormFactor 要素は、アドインのモバイルフォームファクターの設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 64a7681ca23becf42af1ba435aae4d509e6ad1ba
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612228"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="8b176-103">MobileFormFactor 要素</span><span class="sxs-lookup"><span data-stu-id="8b176-103">MobileFormFactor element</span></span>

<span data-ttu-id="8b176-p101">モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。</span><span class="sxs-lookup"><span data-stu-id="8b176-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="8b176-106">各**MobileFormFactor**定義には、 **functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。</span><span class="sxs-lookup"><span data-stu-id="8b176-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="8b176-107">詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="8b176-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="8b176-p103">**MobileFormFactor** 要素は、VersionOverrides のスキーマ 1.1 で定義されています。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。</span><span class="sxs-lookup"><span data-stu-id="8b176-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="8b176-110">子要素</span><span class="sxs-lookup"><span data-stu-id="8b176-110">Child elements</span></span>

| <span data-ttu-id="8b176-111">要素</span><span class="sxs-lookup"><span data-stu-id="8b176-111">Element</span></span>                               | <span data-ttu-id="8b176-112">必須</span><span class="sxs-lookup"><span data-stu-id="8b176-112">Required</span></span> | <span data-ttu-id="8b176-113">説明</span><span class="sxs-lookup"><span data-stu-id="8b176-113">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="8b176-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="8b176-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="8b176-115">はい</span><span class="sxs-lookup"><span data-stu-id="8b176-115">Yes</span></span>      | <span data-ttu-id="8b176-116">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="8b176-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="8b176-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="8b176-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="8b176-118">はい</span><span class="sxs-lookup"><span data-stu-id="8b176-118">Yes</span></span>      | <span data-ttu-id="8b176-119">JavaScript 関数を含むファイルの URL。</span><span class="sxs-lookup"><span data-stu-id="8b176-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="8b176-120">MobileFormFactor の例</span><span class="sxs-lookup"><span data-stu-id="8b176-120">MobileFormFactor example</span></span>

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
