---
title: マニフェストファイルの EquivalentAddin 要素
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059924"
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

[Type](type.md)

## <a name="can-contain"></a>含めることができるもの

[ProgId](progid.md)
[ファイル名](filename.md)

## <a name="remarks"></a>解説

COM アドインを同等のアドインとして指定するには、と`ProgId` `Type`の両方の要素を指定します。 XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Excel アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)