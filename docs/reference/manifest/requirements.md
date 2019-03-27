---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870367"
---
# <a name="requirements-element"></a>Requirements 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。

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

