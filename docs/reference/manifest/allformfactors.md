---
title: マニフェスト ファイルの AllFormFactors 要素
description: すべてのフォーム ファクターについてアドインの設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 674fbe9defa961cb0eef1103cf2dedea0983ffabadc665b172d1f3b15292e987
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088541"
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
