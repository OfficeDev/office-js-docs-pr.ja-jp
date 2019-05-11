---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: b46536886d59692d03976083412a8b8d2e6ae859
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952391"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 要素

デスクトップフォームファクター用のアドインの設定を指定します。 デスクトップフォームファクターには、Office on Windows、Office for Mac、Office Online が含まれています。 このファイルには、[**リソース**] ノードを除くデスクトップフォームファクターのすべてのアドイン情報が含まれています。

各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)       | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](getstarted.md)           | 不要       | Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。 |
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
