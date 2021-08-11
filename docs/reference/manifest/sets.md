---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、アクティブ化Office必要Office JavaScript API の最小セットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a0a7edf6543cc74ac69ee6dc430c0a7497b6911ed43d66ea1082c0d477255948
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095021"
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

