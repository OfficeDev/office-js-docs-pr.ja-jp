---
title: マニフェスト ファイルの Override 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432866"
---
# <a name="override-element"></a>Override 要素

追加ロケールの設定の値を指定する方法を提供します。

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

## <a name="syntax"></a>構文

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a>含まれる場所

|**要素**|
|:-----|
|[CitationText](citationtext.md)|
|[Description](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|

## <a name="attributes"></a>属性

|**属性**|**型**|**必須**|**説明**|
|:-----|:-----|:-----|:-----|
|Locale|string|必須|`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。|
|Value|string|必須|指定のロケールに対して表される設定の値を指定します。|

## <a name="see-also"></a>関連項目

- [Office アドインのローカライズ](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
