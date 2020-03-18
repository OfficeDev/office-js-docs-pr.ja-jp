---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: デスクトップフォームファクター用のアドインの設定を指定します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bfea6900e6b07d8dc7ad5b5256703d873242d88c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718364"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 要素

デスクトップフォームファクター用のアドインの設定を指定します。 デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。 このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。

各 DesktopFormFactor 定義には、 **Functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。 詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)       | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](getstarted.md)           | いいえ       | Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。 |
| [SupportsSharedFolders](supportssharedfolders.md) | いいえ | 代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。 |

## <a name="desktopformfactor-example"></a>DesktopFormFactor の例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
