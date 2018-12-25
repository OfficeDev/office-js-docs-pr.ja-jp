---
title: マニフェスト ファイルの SupportsSharedFolders 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432684"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 要素

代理人のシナリオで Outlook アドインが使用できるかどうかを定義します。 **SupportsSharedFolders** 要素は、[DesktopFormFactor](desktopformfactor.md) の子要素です。 既定では *false* になっています。

> [!IMPORTANT]
> この要素は、[Outlook アドイン要件セットのプレビュー](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)の Exchange Online に対してのみ使用できます。 この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。

**SupportsSharedFolders** 要素の例を次に示します。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
