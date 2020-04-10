---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215041"
---
# <a name="requestedheight-element"></a><span data-ttu-id="1b65a-103">RequestedHeight 要素</span><span class="sxs-lookup"><span data-stu-id="1b65a-103">RequestedHeight element</span></span>

<span data-ttu-id="1b65a-104">コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。</span><span class="sxs-lookup"><span data-stu-id="1b65a-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="1b65a-105">**アドインの種類:** コンテンツ、メール</span><span class="sxs-lookup"><span data-stu-id="1b65a-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b65a-106">構文</span><span class="sxs-lookup"><span data-stu-id="1b65a-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="1b65a-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="1b65a-107">Contained in</span></span>

- <span data-ttu-id="1b65a-108">[DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="1b65a-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="1b65a-109">[DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="1b65a-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="1b65a-110">[ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) では、**DetectedEntity** 拡張点の値は 140 から 450 に、**CustomPane** 拡張点の値は 32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="1b65a-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>
