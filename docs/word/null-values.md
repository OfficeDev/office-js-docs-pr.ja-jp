---
title: Word アドインの Null 値
description: Word アドインで null 値を使用する方法について説明します。
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e21677dafcaaaa7e9e9164ef18c82f49820298d6
ms.sourcegitcommit: 9d930b4c77c342246607aef30479e31fdbdd47f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63353860"
---
# <a name="null-values-in-word-add-ins"></a>Word アドインの Null 値

`null` Word JavaScript API には特別な意味があります。 既定値を表したり、書式設定を使用したりするために使用されます。

## <a name="null-property-values-in-the-response"></a>応答内の null プロパティ値

color などの書式設定プロパティ [には](/javascript/api/word/word.font#word-word-font-color-member) 、指定した `null` 範囲内に異なる値が存在する場合、応答に値が含 [まれます](/javascript/api/word/word.range)。 たとえば、範囲を取得してその `range.font.color` プロパティを読み込む場合:

- 範囲内のすべてのテキストのフォント色が同じ場合は、 `range.font.color` その色を指定します。
- 範囲内に複数のフォントの色がある場合、`range.font.color` は `null` です。
