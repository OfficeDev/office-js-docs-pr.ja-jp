---
title: マニフェスト ファイルの AllFormFactors 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433279"
---
# <a name="allformfactors-element"></a><span data-ttu-id="13e7c-102">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="13e7c-102">AllFormFactors element</span></span>

<span data-ttu-id="13e7c-103">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="13e7c-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="13e7c-104">現在、**AllFormFactors** を使用する機能はカスタム関数のみです。</span><span class="sxs-lookup"><span data-stu-id="13e7c-104">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="13e7c-105">**AllFormFactors** は、カスタム関数を使用するときの必須要素です。</span><span class="sxs-lookup"><span data-stu-id="13e7c-105">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="13e7c-106">子要素</span><span class="sxs-lookup"><span data-stu-id="13e7c-106">Child elements</span></span>

|  <span data-ttu-id="13e7c-107">要素</span><span class="sxs-lookup"><span data-stu-id="13e7c-107">Element</span></span> |  <span data-ttu-id="13e7c-108">必須</span><span class="sxs-lookup"><span data-stu-id="13e7c-108">Required</span></span>  |  <span data-ttu-id="13e7c-109">説明</span><span class="sxs-lookup"><span data-stu-id="13e7c-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="13e7c-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="13e7c-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="13e7c-111">はい</span><span class="sxs-lookup"><span data-stu-id="13e7c-111">Yes</span></span> |  <span data-ttu-id="13e7c-112">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="13e7c-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="13e7c-113">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="13e7c-113">AllFormFactors example</span></span>

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
