---
title: マニフェストファイルの Enabled 要素
description: アドインの起動時にアドインコマンドを無効にするように指定する方法について説明します。
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611569"
---
# <a name="enabled-element"></a>Enabled 要素

アドインを起動するときに、[ボタン](control.md#button-control)または[メニュー](control.md#menu-dropdown-button-controls)コントロールを有効にするかどうかを指定します。 **Enabled**要素は、 [Control](control.md)の子要素です。 省略すると、既定値はに `true` なります。

親コントロールは、プログラムを使用して有効または無効にすることもできます。 詳細については、「[アドイン コマンドを有効または無効にする](../../design/disable-add-in-commands.md)」を参照してください。

## <a name="example"></a>例

```xml
<Enabled>false</Enabled>
```
