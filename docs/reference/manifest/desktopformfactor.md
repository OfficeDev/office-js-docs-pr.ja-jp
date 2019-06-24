---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d1f09203518a38f1568b13e6c1a9c70752697152
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128518"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 要素

デスクトップフォームファクター用のアドインの設定を指定します。 デスクトップフォームファクターには、web、Windows、Mac に Office が含まれています。 このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。

各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)       | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](getstarted.md)           | いいえ       | Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。 |
| [SupportsSharedFolders](supportssharedfolders.md) | いいえ | 代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。<br><br>**重要**: Outlook アドインの代理人アクセスは現在プレビュー段階であるため、この`SupportSharedFolders`要素を使用するアドインは、appsource に発行することも、一元展開によって展開することもできません。 |

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
