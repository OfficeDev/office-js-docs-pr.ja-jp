---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 44675918a4208683f442fe8a6e8f4f906f484571
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611730"
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
