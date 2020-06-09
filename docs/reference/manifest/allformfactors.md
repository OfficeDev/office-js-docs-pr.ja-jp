---
title: マニフェスト ファイルの AllFormFactors 要素
description: すべてのフォーム ファクターについてアドインの設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608797"
---
# <a name="allformfactors-element"></a><span data-ttu-id="5700a-103">AllFormFactors 要素</span><span class="sxs-lookup"><span data-stu-id="5700a-103">AllFormFactors element</span></span>

<span data-ttu-id="5700a-104">すべてのフォーム ファクターについてアドインの設定を指定します。</span><span class="sxs-lookup"><span data-stu-id="5700a-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="5700a-105">現在、**AllFormFactors** を使用する機能はカスタム関数のみです。</span><span class="sxs-lookup"><span data-stu-id="5700a-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="5700a-106">**AllFormFactors** は、カスタム関数を使用するときの必須要素です。</span><span class="sxs-lookup"><span data-stu-id="5700a-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5700a-107">子要素</span><span class="sxs-lookup"><span data-stu-id="5700a-107">Child elements</span></span>

|  <span data-ttu-id="5700a-108">要素</span><span class="sxs-lookup"><span data-stu-id="5700a-108">Element</span></span> |  <span data-ttu-id="5700a-109">必須</span><span class="sxs-lookup"><span data-stu-id="5700a-109">Required</span></span>  |  <span data-ttu-id="5700a-110">説明</span><span class="sxs-lookup"><span data-stu-id="5700a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5700a-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="5700a-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="5700a-112">はい</span><span class="sxs-lookup"><span data-stu-id="5700a-112">Yes</span></span> |  <span data-ttu-id="5700a-113">アドインが機能を公開する場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="5700a-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="5700a-114">AllFormFactors の例</span><span class="sxs-lookup"><span data-stu-id="5700a-114">AllFormFactors example</span></span>

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
