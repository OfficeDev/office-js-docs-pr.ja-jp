---
title: マニフェストファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 425b926901b7325665eeede04263f74e4b854d50
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718287"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 要素

同等の COM アドインまたは XLL の下位互換性を指定します。

**アドインの種類:** 作業ウィンドウ、ユーザー設定関数

## <a name="syntax"></a>構文

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAdd](equivalentaddins.md)

## <a name="must-contain"></a>含める必要があるもの

[型](type.md)

## <a name="can-contain"></a>含めることができるもの

[ProgId](progid.md)
[ファイル名](filename.md)

## <a name="remarks"></a>注釈

COM アドインを同等のアドインとして指定するには、と`ProgId` `Type`の両方の要素を指定します。 XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Excel アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)