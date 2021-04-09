---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652245"
---
# <a name="runtime-element"></a>Runtime 要素

共有 JavaScript ランタイムを使用して、さまざまなコンポーネントすべてが同じランタイムで実行されるアドインを構成します。 要素の [`<Runtimes>`](runtimes.md) 子。

**アドインの種類:** 作業ウィンドウ, メール

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

- [ランタイム](runtimes.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。 |
|  **有効期間**  |  いいえ  | 既定値は `lifetime` is `short` であり、指定する必要はない。 Outlook アドインは値のみを使用 `short` します。 Excel アドインで共有ランタイムを使用する場合は、値を明示的にに設定します `long` 。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)
