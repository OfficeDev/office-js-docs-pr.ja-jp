---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325249"
---
# <a name="requirements-element"></a>Requirements 要素

Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。

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

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

