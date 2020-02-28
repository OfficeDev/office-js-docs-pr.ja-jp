---
title: マニフェスト ファイルの MobileFormFactor 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324842"
---
# <a name="mobileformfactor-element"></a>MobileFormFactor 要素

モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。

各**MobileFormFactor**定義には、 **functionfile**要素と1つ以上の**extensionpoint**要素が含まれています。 詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

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
