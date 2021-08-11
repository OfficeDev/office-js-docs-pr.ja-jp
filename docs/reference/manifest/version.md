---
title: マニフェスト ファイルの Version 要素
description: Version 要素は、アドイン Officeを指定します。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 9641153cbe6fa0284986b8dd286ba2114b32a82894bd5f8d33516e2a56c90be9
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57096331"
---
# <a name="version-element"></a>Version 要素

Office アドインのバージョンを指定します。 バージョン番号は、1、2、3、または 4 パーツ (n、n.n、n.n.n、または n.n.n.n) です。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Version>n[.n.n.n]</Version>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="remarks"></a>注釈

バージョン番号の各部分には、最大 5 桁の数字を指定できます。
