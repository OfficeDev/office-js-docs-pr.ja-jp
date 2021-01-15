---
title: マニフェスト ファイルの Enabled 要素
description: アドインの起動時にアドイン コマンドを無効に指定する方法について説明します。
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771398"
---
# <a name="enabled-element"></a>Enabled 要素

アドインの起動時 [にボタン](control.md#button-control) コントロールまたは [メニュー](control.md#menu-dropdown-button-controls) コントロールを有効にするかどうかを指定します。 **Enabled 要素** は [、Control](control.md)の子要素です。 省略すると、既定値は `true` .

この要素は Excel でのみ有効です。つまり `Name` [、Host](host.md) 要素の属性が "Workbook" の場合です。

親コントロールは、プログラムで有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
