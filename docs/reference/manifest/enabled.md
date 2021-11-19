---
title: マニフェスト ファイル内の Enabled 要素
description: アドインの起動時にアドイン コマンドが無効になっていることを指定する方法について説明します。
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4c0107daaf73aee6ba116553a8d01250e9c7d981
ms.sourcegitcommit: 997a20f9fb011b96a50ceb04a4b9943d92d6ecf4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/19/2021
ms.locfileid: "61081436"
---
# <a name="enabled-element"></a>Enabled 要素

アドインの起動時に [Button](control.md#button-control) コントロールまたは [Menu](control.md#menu-dropdown-button-controls) コントロールを有効にするかどうかを指定します。 **Enabled 要素** は、Control の子要素 [です](control.md)。 省略すると、既定値は `true` .

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

この要素は、Host 要素 `Name` [Excel"Workbook"](host.md)の場合にのみ有効です。

親コントロールは、プログラムで有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
