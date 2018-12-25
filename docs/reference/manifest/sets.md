---
title: マニフェスト ファイルの Sets 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: b7e78ae05f8409f38c885a1d6a328347d00d0df1
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433657"
---
# <a name="sets-element"></a>Sets 要素

Office アドインをアクティブにするために必要な JavaScript API for Office の最小限のサブセットを指定します。

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

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|DefaultMinVersion|文字列|省略可能|すべての子の **Set** 要素に対して、既定の [MinVersion](set.md) 属性値を指定します。既定値は "1.1" です。|

## <a name="remarks"></a>解説

利用できる要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。

