---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメールアドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5f4c3ca1ff39cc3150249fbc824b0db76f6b8a85
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215041"
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
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) では、**DetectedEntity** 拡張点の値は 140 から 450 に、**CustomPane** 拡張点の値は 32 から 450 にすることが可能
