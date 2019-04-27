---
title: マニフェストファイルの EquivalentAddin 要素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356896"
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

[ProgID](progid.md)
[ファイル名](filename.md)

## <a name="remarks"></a>注釈

COM アドインを同等のアドインとして指定するには、と`ProgID` `Type`の両方の要素を指定します。 XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性があるようにする](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [既存の COM アドインと互換性のある Office アドインを作成する](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)