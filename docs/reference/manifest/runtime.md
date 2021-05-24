---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590914"
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

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [Override](override.md) | なし | **Outlook**: LaunchEvent 拡張ポイント ハンドラーにデスクトップで必要Outlook JavaScript ファイルの URL [の場所を指定](../../reference/manifest/extensionpoint.md#launchevent)します。 **重要**: 現時点では、定義できる要素は 1 つで `<Override>` 、型である必要があります `javascript` 。|

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid**  |  必要  | アドインの HTML ページの URL の場所を指定します。 32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。 |
|  **有効期間**  |  いいえ  | 既定値は `lifetime` is `short` であり、指定する必要はない。 Outlookは値のみを使用 `short` します。 アドインで共有ランタイムを使用する場合Excelに値を明示的に設定します `long` 。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのOutlook用にアドインを構成する](../../outlook/autolaunch.md)
