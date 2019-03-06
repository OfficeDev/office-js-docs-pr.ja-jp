---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413616"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> Outlook アドインの代理人アクセスは現在プレビュー段階であり、Exchange Online に対して実行されるクライアントでのみサポートされています。 この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。

**SupportsSharedFolders** 要素の例を次に示します。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
