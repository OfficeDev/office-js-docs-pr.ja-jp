---
title: マニフェストファイルの Type 要素
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628229"
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

- COM: 対応するアドインが COM アドインであることを指定します。
- XLL: 対応するアドインが Excel XLL であることを指定します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Excel アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)