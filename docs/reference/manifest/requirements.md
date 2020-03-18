---
title: マニフェスト ファイルの Requirements 要素
description: 要件要素は、Office アドインをアクティブにするために必要な最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a3f41a763ec820a6c766e6a32b26e55ad34996f7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720450"
---
# <a name="requirements-element"></a>Requirements 要素

Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

|**Element**|**コンテンツ**|**メール**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[メソッド](methods.md)|x||x|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。
