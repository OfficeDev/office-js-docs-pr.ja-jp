---
title: マニフェスト ファイルの Version 要素
description: Version 要素は、アドインOfficeバージョンを指定します。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173935"
---
# <a name="version-element"></a>Version 要素

Office アドインのバージョンを指定します。 バージョン番号は、1、2、3、または 4 つの部分 (つまり、n、n.n、n.n.n、または n.n.n.n) です。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

バージョン番号の各部分には、最大 5 桁の数字を指定できます。
