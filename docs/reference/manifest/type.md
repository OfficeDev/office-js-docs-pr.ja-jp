---
title: マニフェストファイルの Type 要素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356905"
---
# <a name="type-element"></a>Type 要素

対応するアドインが COM addin または XLL であるかどうかを指定します。

**アドインの種類:** 作業ウィンドウ、ユーザー設定関数

## <a name="syntax"></a>構文

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAdd](equivalentaddin.md)

## <a name="add-in-type-values"></a>アドインの種類の値

`Type`要素には、次のいずれかの値を指定する必要があります。

- com: 対応するアドインが COM アドインであることを指定します。
- xll: 対応するアドインが Excel XLL であることを指定します。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性があるようにする](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Office アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)