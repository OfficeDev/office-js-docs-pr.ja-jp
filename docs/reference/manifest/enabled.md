---
title: マニフェスト ファイルの Enabled 要素
description: アドインの起動時にアドイン コマンドを無効に指定する方法について説明します。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771398"
---
# <a name="enabled-element"></a><span data-ttu-id="d2355-103">Enabled 要素</span><span class="sxs-lookup"><span data-stu-id="d2355-103">Enabled element</span></span>

<span data-ttu-id="d2355-104">アドインの起動時 [にボタン](control.md#button-control) コントロールまたは [メニュー](control.md#menu-dropdown-button-controls) コントロールを有効にするかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="d2355-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="d2355-105">**Enabled 要素** は [、Control](control.md)の子要素です。</span><span class="sxs-lookup"><span data-stu-id="d2355-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="d2355-106">省略すると、既定値は `true` .</span><span class="sxs-lookup"><span data-stu-id="d2355-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="d2355-107">この要素は Excel でのみ有効です。つまり `Name` [、Host](host.md) 要素の属性が "Workbook" の場合です。</span><span class="sxs-lookup"><span data-stu-id="d2355-107">This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".</span></span>

<span data-ttu-id="d2355-108">親コントロールは、プログラムで有効または無効にすることもできます。</span><span class="sxs-lookup"><span data-stu-id="d2355-108">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="d2355-109">詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d2355-109">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="d2355-110">例</span><span class="sxs-lookup"><span data-stu-id="d2355-110">Example</span></span>

```xml
<Enabled>false</Enabled>
```
