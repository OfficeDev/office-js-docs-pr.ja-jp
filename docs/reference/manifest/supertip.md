---
title: マニフェスト ファイルの Supertip 要素
description: Supertip 要素は、リッチ ヒント (タイトルと説明の両方) を定義します。
ms.date: 05/07/2019
localization_priority: Normal
ms.openlocfilehash: 79120cc72aa4804eaaa2330d9298f6521a13552d325d9134814581402ace8210
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093262"
---
# <a name="supertip"></a>Supertip

豊富なヒント (タイトルと説明の両方) を定義します。[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの両方で使用されます。

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [Title](#title) | はい | ヒントのテキストです。 |
| [説明](#description) | はい | ヒントの説明です。<br>**注**: (Outlook) Windows Mac クライアントだけがサポートされています。 |

### <a name="title"></a>タイトル

必ず指定します。 ヒントのテキストです。 **resid 属性** は 32 文字以内で、Resources 要素の **ShortStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

### <a name="description"></a>説明

必ず指定します。 ヒントの記述です。 **resid 属性** は 32 文字以内で、Resources 要素の **LongStrings** 要素の **String** 要素の **id** 属性の値に設定 [する必要](resources.md)があります。

> [!NOTE]
> このOutlook、Windows Mac クライアントだけが Description 要素を **サポート** します。

## <a name="example"></a>例

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
