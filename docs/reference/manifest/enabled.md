---
title: マニフェスト ファイル内の Enabled 要素
description: アドインの起動時にアドイン コマンドが無効になっていることを指定する方法について説明します。
ms.date: 03/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 75a2664143e29c86a05aaf039b0ea7bce659cef9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744757"
---
# <a name="enabled-element"></a>Enabled 要素

アドインの起動時に [Button コントロール](control-button.md) または [Menu](control-menu.md) コントロールを有効にするかどうかを指定します。 **Enabled 要素** は、Control の子要素 [です](control.md)。 省略すると、既定値は `true`.

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [RibbonApi 1.0](../requirement-sets/ribbon-api-requirement-sets.md)

この要素は、Excel、PowerPoint、および Word `Name` でのみ有効です。つまり、[Host](host.md) 要素の属性が "Workbook"、"Presentation"、または "Document" の場合です。

親コントロールは、プログラムで有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
