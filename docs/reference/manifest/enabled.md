---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566203"
---
# <a name="enabled-element"></a><span data-ttu-id="2624d-103">Enabled 要素</span><span class="sxs-lookup"><span data-stu-id="2624d-103">Enabled element</span></span>

<span data-ttu-id="2624d-104">アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="2624d-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="2624d-105">**Enabled**要素は、 [Control](control.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="2624d-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="2624d-106">省略すると、既定値は`true`になります。</span><span class="sxs-lookup"><span data-stu-id="2624d-106">If it is omitted, the default is `true`.</span></span> 

<span data-ttu-id="2624d-107">親コントロールは、プログラムを使用して有効または無効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="2624d-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="2624d-108">詳細については、「[アドインコマンドを有効または無効](/office/dev/add-ins/design/disable-add-in-commands)にする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2624d-108">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

## <a name="example"></a><span data-ttu-id="2624d-109">例</span><span class="sxs-lookup"><span data-stu-id="2624d-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```

