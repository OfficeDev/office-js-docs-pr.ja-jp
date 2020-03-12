---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596901"
---
# <a name="enabled-element"></a><span data-ttu-id="59e0a-103">Enabled 要素</span><span class="sxs-lookup"><span data-stu-id="59e0a-103">Enabled element</span></span>

<span data-ttu-id="59e0a-104">アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="59e0a-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="59e0a-105">**Enabled**要素は、 [Control](control.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="59e0a-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="59e0a-106">省略すると、既定値は`true`になります。</span><span class="sxs-lookup"><span data-stu-id="59e0a-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="59e0a-107">親コントロールは、プログラムを使用して有効または無効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="59e0a-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="59e0a-108">詳細については、「[アドインコマンドを有効または無効](../../design/disable-add-in-commands.md)にする」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="59e0a-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="59e0a-109">例</span><span class="sxs-lookup"><span data-stu-id="59e0a-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
