---
title: マニフェスト ファイルの Type 要素
description: Type 要素は、同等のアドインが COM アドインか XLL かを指定します。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: ca6fa7183727870593dd3e726abc72fdc0d6f0b518fdb8451ec80c6b590f8c83
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092479"
---
# <a name="type-element"></a>Type 要素

同等のアドインが COM アドインか XLL かを指定します。

**アドインの種類:** 作業ウィンドウ、カスタム関数

## <a name="syntax"></a>構文

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAddin](equivalentaddin.md)

## <a name="add-in-type-values"></a>アドインの型の値

要素には、次のいずれかの値を指定する必要 `Type` があります。

- COM: COM アドインと同等のアドインを指定します。
- XLL: 同等のアドインが XLL のExcelします。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Office アドインを既存の COM アドインと互換できるようにする](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)