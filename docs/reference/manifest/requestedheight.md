---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: fa40043e6192e1304e67f1f96f770898b230036c
ms.sourcegitcommit: b634bfe9a946fbd95754e87f070a904ed57586ff
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/15/2020
ms.locfileid: "44253615"
---
# <a name="requestedheight-element"></a>RequestedHeight 要素

コンテンツ アドインまたはメール アドインの初期高さ (ピクセル単位) を指定します。

**アドインの種類:** コンテンツ、メール

## <a name="syntax"></a>構文

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a>含まれる場所

- [DefaultSettings](defaultsettings.md) (コンテンツ アドイン) の値は、32 から 1000 にすることが可能
- [DesktopSettings](desktopsettings.md) と [TabletSettings](tabletsettings.md) (メール アドイン) の値は、32 から 450 にすることが可能
- [Extensionpoint](extensionpoint.md) (コンテキストメールアドイン) は、 **DetectedEntity**拡張点の場合は140と450、 [ **custompane**拡張点](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)の場合は32と450の間で指定できます (非推奨)。
