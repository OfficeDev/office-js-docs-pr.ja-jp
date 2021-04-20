---
title: マニフェスト ファイルの Supertip 要素
description: Supertip 要素は、豊富なヒント (タイトルと説明の両方) を定義します。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 5e8b3850d99f6791726b1b2f0545c5fb4b52c554
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771299"
---
# <a name="supertip"></a>Supertip

豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [Title](#title) | はい | ヒントのテキストです。 |
| [説明](#description) | はい | ヒントの説明です。<br>**注**: (Outlook) Windows および Mac クライアントだけがサポートされています。 |

### <a name="title"></a>タイトル

必ず指定します。 ヒントのテキストです。 **resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

### <a name="description"></a>説明

必ず指定します。 ヒントの記述です。 **resid 属性** は 32 文字以内で [、Resources](resources.md)要素の **LongStrings** 要素の **String** 要素の **id** 属性の値に設定する必要があります。

> [!NOTE]
> Outlook では、Description 要素をサポートしているのは Windows クライアントと Mac **クライアント** のみです。

## <a name="example"></a>例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
