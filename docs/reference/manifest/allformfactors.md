---
title: マニフェスト ファイルの AllFormFactors 要素
description: すべてのフォーム ファクターについてアドインの設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f1285f92b5eb89993e7fcfe79aab2325b86aca3d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720716"
---
# <a name="allformfactors-element"></a><span data-ttu-id="b5118-103">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="b5118-103">AllFormFactors element</span></span>

<span data-ttu-id="b5118-104">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="b5118-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="b5118-105">現在、**AllFormFactors** を使用する機能はカスタム関数のみです。</span><span class="sxs-lookup"><span data-stu-id="b5118-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="b5118-106">**AllFormFactors** は、カスタム関数を使用するときの必須要素です。</span><span class="sxs-lookup"><span data-stu-id="b5118-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b5118-107">子要素</span><span class="sxs-lookup"><span data-stu-id="b5118-107">Child elements</span></span>

|  <span data-ttu-id="b5118-108">要素</span><span class="sxs-lookup"><span data-stu-id="b5118-108">Element</span></span> |  <span data-ttu-id="b5118-109">必須</span><span class="sxs-lookup"><span data-stu-id="b5118-109">Required</span></span>  |  <span data-ttu-id="b5118-110">説明</span><span class="sxs-lookup"><span data-stu-id="b5118-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b5118-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b5118-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="b5118-112">はい</span><span class="sxs-lookup"><span data-stu-id="b5118-112">Yes</span></span> |  <span data-ttu-id="b5118-113">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="b5118-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="b5118-114">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="b5118-114">AllFormFactors example</span></span>

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
