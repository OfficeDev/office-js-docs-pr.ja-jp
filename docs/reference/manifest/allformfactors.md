---
title: マニフェスト ファイルの AllFormFactors 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450738"
---
# <a name="allformfactors-element"></a><span data-ttu-id="7767f-102">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="7767f-102">AllFormFactors element</span></span>

<span data-ttu-id="7767f-103">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="7767f-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="7767f-104">現在、**AllFormFactors** を使用する機能はカスタム関数のみです。</span><span class="sxs-lookup"><span data-stu-id="7767f-104">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="7767f-105">**AllFormFactors** は、カスタム関数を使用するときの必須要素です。</span><span class="sxs-lookup"><span data-stu-id="7767f-105">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7767f-106">子要素</span><span class="sxs-lookup"><span data-stu-id="7767f-106">Child elements</span></span>

|  <span data-ttu-id="7767f-107">要素</span><span class="sxs-lookup"><span data-stu-id="7767f-107">Element</span></span> |  <span data-ttu-id="7767f-108">必須</span><span class="sxs-lookup"><span data-stu-id="7767f-108">Required</span></span>  |  <span data-ttu-id="7767f-109">説明</span><span class="sxs-lookup"><span data-stu-id="7767f-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7767f-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="7767f-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="7767f-111">はい</span><span class="sxs-lookup"><span data-stu-id="7767f-111">Yes</span></span> |  <span data-ttu-id="7767f-112">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="7767f-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="7767f-113">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="7767f-113">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
