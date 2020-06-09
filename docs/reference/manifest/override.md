---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、追加のロケールの設定値を指定できます。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aa5d023169389670d15e36f8bee4445529d84711
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611506"
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

- [Office アドインのローカライズ](../../develop/localization.md)
