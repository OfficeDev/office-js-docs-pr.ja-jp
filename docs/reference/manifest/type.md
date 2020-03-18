---
title: マニフェストファイルの Type 要素
description: Type 要素は、対応するアドインが COM アドインまたは XLL であるかどうかを指定します。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720317"
---
# <a name="type-element"></a>Type 要素

対応するアドインが COM アドインまたは XLL であるかどうかを指定します。

**アドインの種類:** 作業ウィンドウ、ユーザー設定関数

## <a name="syntax"></a>構文

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAdd](equivalentaddin.md)

## <a name="add-in-type-values"></a>アドインの種類の値

`Type`要素には、次のいずれかの値を指定する必要があります。

- COM: 対応するアドインが COM アドインであることを指定します。
- XLL: 対応するアドインが Excel XLL であることを指定します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Excel アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)