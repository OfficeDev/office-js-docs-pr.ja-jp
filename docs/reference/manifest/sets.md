---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、office アドインをアクティブにするために必要な最低限の Office JavaScript API のセットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641425"
---
# <a name="sets-element"></a>Sets 要素

Office アドインをアクティブにするために必要な最低限の Office JavaScript API のサブセットを指定します。

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
|DefaultMinVersion|文字列|省略可能|すべての子[セット](set.md)要素の既定の**MinVersion**属性値を指定します。 既定値は "1.1" です。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。

