---
title: マニフェスト ファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 03/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: e0842a0a5663d6221e854881aef28004ad2a8705
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153047"
---
# <a name="equivalentaddin-element"></a>EquivalentAddin 要素

同等の COM アドインまたは XLL の下位互換性を指定します。

**アドインの種類:** 作業ウィンドウ、カスタム関数

## <a name="syntax"></a>構文

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a>含まれる場所

[EquivalentAddins](equivalentaddins.md)

## <a name="must-contain"></a>含める必要があるもの

[Type](type.md)

## <a name="can-contain"></a>含めることができるもの

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>解説

COM アドインを同等のアドインとして指定するには、要素と要素の両方を `ProgId` 指定 `Type` します。 XLL を同等のアドインとして指定するには、要素と要素の両方を `FileName` 指定 `Type` します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Office アドインを既存の COM アドインと互換できるようにする](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)