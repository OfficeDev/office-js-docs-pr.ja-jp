---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566203"
---
# <a name="enabled-element"></a>Enabled 要素

アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。 **Enabled**要素は、 [Control](control.md)の子要素です。 省略すると、既定値は`true`になります。 

親コントロールは、プログラムを使用して有効または無効にすることもできます。 詳細については、「[アドインコマンドを有効または無効](/office/dev/add-ins/design/disable-add-in-commands)にする」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```

