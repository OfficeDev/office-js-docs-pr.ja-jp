---
title: マニフェスト ファイル内の Enabled 要素
description: アドインの起動時にアドイン コマンドが無効になっていることを指定する方法について説明します。
ms.date: 01/04/2021
ms.localizationpriority: medium
ms.openlocfilehash: a14385f7114eb3d35845b5d9873bdd718b46c0e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154003"
---
# <a name="enabled-element"></a>Enabled 要素

アドインの起動時に [Button](control.md#button-control) コントロールまたは [Menu](control.md#menu-dropdown-button-controls) コントロールを有効にするかどうかを指定します。 **Enabled 要素** は、Control の子要素 [です](control.md)。 省略すると、既定値は `true` .

この要素は、この要素のExcel。つまり、Host 要素の `Name` 属性が["Workbook"](host.md)の場合です。

親コントロールは、プログラムで有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
