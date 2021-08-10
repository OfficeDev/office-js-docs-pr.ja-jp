---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメール アドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 0e5f9de909d32622ac244ff4118c8a3192abf2ff0fe89ed81a6188ddcb265549
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092987"
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
- [ExtensionPoint](extensionpoint.md) (コンテキスト メール アドイン) の値は **、DetectedEntity** 拡張ポイントの場合は 140 ~ 450、CustomPane 拡張ポイントの場合は 32 ~ 450 です (非推奨 [)](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)
