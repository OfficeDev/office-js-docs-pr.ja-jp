---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、アクティブ化Office必要Office JavaScript API の最小セットを指定します。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 38707ec78a79e9104dd21f9fa5ceab8c6fbd2c79
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154447"
---
# <a name="sets-element"></a>Sets 要素

アクティブ化するために必要Office JavaScript API の最小Officeサブセットを指定します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a>含まれる場所

[Requirements](requirements.md)

## <a name="can-contain"></a>含めることができるもの

[Set](set.md)

## <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|文字列|省略可能|すべての子 Set 要素 **の既定の MinVersion** 属性値を [指定](set.md) します。 既定値は "1.1" です。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

**Set** 要素の **MinVersion** 属性と Sets 要素 **の DefaultMinVersion** 属性の詳細については、「Set [the Requirements element in the manifest」を参照してください](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)。

