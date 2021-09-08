---
title: マニフェスト ファイルの Version 要素
description: Version 要素は、アドイン Officeを指定します。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 48a2be94d95ece597e47468bb18db2a7962a51e9
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937027"
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
