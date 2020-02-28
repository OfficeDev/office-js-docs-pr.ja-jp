---
title: マニフェスト ファイルの Supertip 要素
description: ''
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: ab280ec550a58f85082c36a24f5f7c3b4112a214
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325235"
---
# <a name="supertip"></a>Supertip

豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [Title](#title) | はい | ヒントのテキストです。 |
| [説明](#description) | はい | ヒントの説明です。<br>**注**: (Outlook) は、Windows および Mac クライアントのみがサポートされています。 |

### <a name="title"></a>Title

必須です。 ヒントのテキスト。 **Resid**属性は、 [Resources](resources.md)要素の Short **strings**要素の**String**要素の**id**属性の値に設定する必要があります。

### <a name="description"></a>説明

必ず指定します。 ヒントの記述です。 **Resid**属性は、 [Resources](resources.md)要素の**longstrings**要素の**String**要素の**id**属性の値に設定する必要があります。

> [!NOTE]
> Outlook の場合、Windows と Mac のクライアントのみが**Description**要素をサポートしています。

## <a name="example"></a>例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
