---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: fa40043e6192e1304e67f1f96f770898b230036c
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253615"
---
# <a name="requestedheight-element"></a><span data-ttu-id="3c563-103">RequestedHeight 要素</span><span class="sxs-lookup"><span data-stu-id="3c563-103">RequestedHeight element</span></span>

<span data-ttu-id="3c563-104">コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。</span><span class="sxs-lookup"><span data-stu-id="3c563-104">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span>

<span data-ttu-id="3c563-105">**アドインの種類:** コンテンツ、メール</span><span class="sxs-lookup"><span data-stu-id="3c563-105">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3c563-106">構文</span><span class="sxs-lookup"><span data-stu-id="3c563-106">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="3c563-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="3c563-107">Contained in</span></span>

- <span data-ttu-id="3c563-108">[DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="3c563-108">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="3c563-109">[DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能</span><span class="sxs-lookup"><span data-stu-id="3c563-109">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="3c563-110">[Extensionpoint](extensionpoint.md) (コンテキストメールアドイン) は、 **DetectedEntity**拡張点の場合は140と450、 [ **custompane**拡張点](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)の場合は32と450の間で指定できます (非推奨)。</span><span class="sxs-lookup"><span data-stu-id="3c563-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the [**CustomPane** extension point (deprecated)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)</span></span>
