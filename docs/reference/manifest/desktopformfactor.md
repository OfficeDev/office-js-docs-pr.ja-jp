---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433741"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 要素

デスクトップ フォーム ファクターについてアドインの設定を指定します。デスクトップ フォーム ファクターには、Office for Windows、Office for Mac、Office Online が含まれています。**Resources** ノードを除くデスクトップ フォーム ファクターのアドイン情報をすべて含みます。

各 DesktopFormFactor の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)       | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](getstarted.md)           | いいえ       | Word、Excel、または PowerPoint のホストにアドインをインストールするときに表示される吹き出しを定義します。 |
| [SupportsSharedFolders](supportssharedfolders.md) | いいえ | 代理人のシナリオで Outlook アドインを使用できるかどうかを定義し、既定では *false* に設定します。<br><br>**重要事項**: この要素は、Outlook アドイン要件セットのプレビューの Exchange Online に対してのみ使用できます。 この要素を使用するアドインは、AppSource に発行したり、一元展開によって展開したりすることはできません。 |

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
