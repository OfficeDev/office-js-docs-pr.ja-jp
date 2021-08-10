---
title: マニフェスト ファイル内の Enabled 要素
description: アドインの起動時にアドイン コマンドが無効になっていることを指定する方法について説明します。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: 54d28839a274ff41bab0b1e2cdd2d169e76c5815095950dec67ce2564eade601
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093907"
---
# <a name="enabled-element"></a>Enabled 要素

アドインの起動時に [Button](control.md#button-control) コントロールまたは [Menu](control.md#menu-dropdown-button-controls) コントロールを有効にするかどうかを指定します。 **Enabled 要素** は、Control の子要素 [です](control.md)。 省略すると、既定値は `true` .

この要素は、この要素のExcel。つまり、Host 要素の `Name` 属性が["Workbook"](host.md)の場合です。

親コントロールは、プログラムで有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
