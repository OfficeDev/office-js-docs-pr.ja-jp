---
title: マニフェスト ファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 6404ad561a14a451e4685cc23be930b7ba612e85d1b37e78aa45f9366becf3bc
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57085763"
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

[型](type.md)

## <a name="can-contain"></a>含めることができるもの

[ProgId](progid.md) 
[FileName](filename.md)

## <a name="remarks"></a>解説

COM アドインを同等のアドインとして指定するには、要素と要素の両方を `ProgId` 指定 `Type` します。 XLL を同等のアドインとして指定するには、要素と要素の両方を `FileName` 指定 `Type` します。

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Office アドインを既存の COM アドインと互換できるようにする](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)