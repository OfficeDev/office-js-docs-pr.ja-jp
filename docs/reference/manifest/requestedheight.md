---
title: マニフェスト ファイルの RequestedHeight 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e175d9012bb2f2a42fd466c35e5e28ade967d6f2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450528"
---
# <a name="requestedheight-element"></a><span data-ttu-id="cfa9a-102">RequestedHeight 要素</span><span class="sxs-lookup"><span data-stu-id="cfa9a-102">RequestedHeight element</span></span>

<span data-ttu-id="cfa9a-103">コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。</span><span class="sxs-lookup"><span data-stu-id="cfa9a-103">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="cfa9a-104">**アドインの種類:** コンテンツ、メール</span><span class="sxs-lookup"><span data-stu-id="cfa9a-104">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cfa9a-105">構文</span><span class="sxs-lookup"><span data-stu-id="cfa9a-105">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="cfa9a-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="cfa9a-106">Contained in</span></span>

- <span data-ttu-id="cfa9a-107">[DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="cfa9a-107">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="cfa9a-108">[DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="cfa9a-108">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="cfa9a-109">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) では、**DetectedEntity** 拡張点の値は 140 から 450 に、**CustomPane** 拡張点の値は 32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="cfa9a-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
