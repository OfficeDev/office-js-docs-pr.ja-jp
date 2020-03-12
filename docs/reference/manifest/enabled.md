---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596901"
---
# <a name="enabled-element"></a>Enabled 要素

アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。 **Enabled**要素は、 [Control](control.md)の子要素です。 省略すると、既定値は`true`になります。

親コントロールは、プログラムを使用して有効または無効にすることもできます。 詳細については、「[アドインコマンドを有効または無効](../../design/disable-add-in-commands.md)にする」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
