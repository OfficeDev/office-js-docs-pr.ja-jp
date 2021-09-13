---
title: マニフェスト ファイルの Requirements 要素
description: Requirements 要素は、アクティブ化するためにアドインに必要Office最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 3a5a393485094b5cc830b5120c3abd8c211eff1e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154075"
---
# <a name="requirements-element"></a>Requirements 要素

アドインがアクティブ化する必要Office JavaScript API 要件 (要件[セット](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets)またはメソッド) の最小セットOffice指定します。

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

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[メソッド](methods.md)|x||x|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。
