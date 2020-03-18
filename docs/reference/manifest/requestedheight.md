---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 853d12baf290167f3e6a635201e8b5d1d0e35a51
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720457"
---
# <a name="requestedheight-element"></a><span data-ttu-id="6ff80-103">RequestedHeight 要素</span><span class="sxs-lookup"><span data-stu-id="6ff80-103">RequestedHeight element</span></span>

<span data-ttu-id="6ff80-104">コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。</span><span class="sxs-lookup"><span data-stu-id="6ff80-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="6ff80-105">**アドインの種類:** コンテンツ、メール</span><span class="sxs-lookup"><span data-stu-id="6ff80-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6ff80-106">構文</span><span class="sxs-lookup"><span data-stu-id="6ff80-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="6ff80-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="6ff80-107">Contained in</span></span>

- <span data-ttu-id="6ff80-108">[DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="6ff80-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="6ff80-109">[DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="6ff80-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="6ff80-110">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) では、**DetectedEntity** 拡張点の値は 140 から 450 に、**CustomPane** 拡張点の値は 32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="6ff80-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
