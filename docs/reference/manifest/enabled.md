---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611569"
---
# <a name="enabled-element"></a><span data-ttu-id="d6387-103">Enabled 要素</span><span class="sxs-lookup"><span data-stu-id="d6387-103">Enabled element</span></span>

<span data-ttu-id="d6387-104">アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="d6387-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="d6387-105">**Enabled**要素は、 [Control](control.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="d6387-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="d6387-106">省略すると、既定値はに `true` なります。</span><span class="sxs-lookup"><span data-stu-id="d6387-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="d6387-107">親コントロールは、プログラムを使用して有効または無効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="d6387-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="d6387-108">詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d6387-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="d6387-109">例</span><span class="sxs-lookup"><span data-stu-id="d6387-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
