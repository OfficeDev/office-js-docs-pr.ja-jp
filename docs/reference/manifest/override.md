---
title: マニフェスト ファイルの Override 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870059"
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
|[説明](description.md)|
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

- [Office アドインのローカライズ](/office/dev/add-ins/develop/localization)
    
