---
title: マニフェストファイルの Type 要素
description: Type 要素は、対応するアドインが COM アドインまたは XLL であるかどうかを指定します。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: b59f903af39facd7543e7384189817d5365cf8c9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604560"
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

要素には、次のいずれかの値を指定する必要があり `Type` ます。

- COM: 対応するアドインが COM アドインであることを指定します。
- XLL: 対応するアドインが Excel XLL であることを指定します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Excel アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)