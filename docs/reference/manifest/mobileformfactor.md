---
title: マニフェスト ファイルの MobileFormFactor 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: f0a68c7127f7872207a58ed252def7a2977c33ed
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433695"
---
# <a name="mobileformfactor-element"></a>MobileFormFactor 要素

モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。

各 **MobileFormFactor** の定義には、**FunctionFile** 要素と、1 つ以上の **ExtensionPoint** 要素が含まれます。詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

**MobileFormFactor** 要素は、VersionOverrides のスキーマ 1.1 で定義されています。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)     | はい      | JavaScript 関数を含むファイルの URL。|

## <a name="mobileformfactor-example"></a>MobileFormFactor の例

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
