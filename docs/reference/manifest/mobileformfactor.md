---
title: マニフェスト ファイルの MobileFormFactor 要素
description: MobileFormFactor 要素は、アドインのモバイル フォーム ファクター設定を指定します。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: c0b701064f3a8c0244dc62226899c4317b9075c4
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990678"
---
# <a name="mobileformfactor-element"></a>MobileFormFactor 要素

モバイル フォーム ファクターについてアドインの設定を指定します。**Resources** ノードを除くモバイル フォーム ファクターのアドイン情報をすべて含みます。

各 **MobileFormFactor 定義には****、FunctionFile** 要素と 1 つ以上の **ExtensionPoint 要素が含** まれています。 詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

**MobileFormFactor** 要素は、VersionOverrides のスキーマ 1.1 で定義されています。これを収容している [VersionOverrides](versionoverrides.md) 要素は、`xsi:type` 属性の値が `VersionOverridesV1_1` になっている必要があります。

**アドインの種類:** メール

## <a name="child-elements"></a>子要素

| 要素                             | 必須 | 説明  |
|:------------------------------------|:--------:|:-------------|
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
