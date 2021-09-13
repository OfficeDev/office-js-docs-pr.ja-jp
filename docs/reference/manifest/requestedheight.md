---
title: マニフェスト ファイルの RequestedHeight 要素
description: RequestedHeight 要素は、コンテンツまたはメール アドインの初期の高さ (ピクセル単位) を指定します。
ms.date: 05/14/2020
ms.localizationpriority: medium
ms.openlocfilehash: e0589e81e8905c4fc8c7a8e50ec7c14038035677
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151327"
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
