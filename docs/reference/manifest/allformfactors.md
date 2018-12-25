---
title: マニフェスト ファイルの AllFormFactors 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433279"
---
# <a name="allformfactors-element"></a>AllFormFactors 要素

すべてのフォーム ファクターについてアドインの設定を指定します。 現在、**AllFormFactors** を使用する機能はカスタム関数のみです。 **AllFormFactors** は、カスタム関数を使用するときの必須要素です。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  はい |  アドインが機能を公開する場所を定義します。 |

## <a name="allformfactors-example"></a>AllFormFactors の例

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
