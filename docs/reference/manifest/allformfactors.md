---
title: マニフェスト ファイルの AllFormFactors 要素
description: すべてのフォーム ファクターについてアドインの設定を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608797"
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
