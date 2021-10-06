---
title: マニフェスト ファイルの DesktopFormFactor 要素
description: デスクトップ フォーム ファクターのアドインの設定を指定します。
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 52c9a029e3f43e9b7d5416455eb99ef3de4dae7a
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138731"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 要素

デスクトップ フォーム ファクターのアドインの設定を指定します。 デスクトップ フォーム ファクターには、Office on the web、Windows Mac が含まれます。 Resources ノードを除く、デスクトップ フォーム ファクターのすべてのアドイン情報が **含** まれる。

各 DesktopFormFactor 定義には **、FunctionFile** 要素と 1 つ以上の **ExtensionPoint 要素が含** まれています。 詳細については、「[FunctionFile 要素](functionfile.md)」と「[ExtensionPoint 要素](extensionpoint.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ, メール

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.0
- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="child-elements"></a>子要素

| 要素                               | 必須 | 説明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | はい      | アドインが機能を公開する場所を定義します。 |
| [FunctionFile](functionfile.md)       | はい      | JavaScript 関数を含むファイルの URL。|
| [GetStarted](getstarted.md)           | いいえ       | Word、Excel、またはアドインにアドインをインストールするときに表示される吹き出しをPowerPoint。 省略すると、吹き出しは [DisplayName](displayname.md) 要素と Description 要素の値 [を](description.md) 使用します。 |
| [SupportsSharedFolders](supportssharedfolders.md) | いいえ | 共有メールボックス (プレビュー Outlook共有フォルダー (つまり、代理アクセス) のシナリオで、アドインを使用できるかどうかを定義します。 既定では *false に* 設定されます。 |

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
