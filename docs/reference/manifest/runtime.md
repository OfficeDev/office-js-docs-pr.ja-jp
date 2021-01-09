---
title: マニフェスト ファイル内のランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するアドインを構成します。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789185"
---
# <a name="runtime-element-preview"></a>Runtime 要素 (プレビュー)

さまざまなコンポーネントが同じランタイムで実行されるのを確認するために、共有 JavaScript ランタイムを使用するアドインを構成します。 要素の [`<Runtimes>`](runtimes.md) 子。

Excel では、この要素により、リボン、作業ウィンドウ、およびカスタム関数で同じランタイムを使用できます。 詳細については、「共有 [JavaScript ランタイムを使用するために Excel](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する」を参照してください。

Outlook では、この要素によってイベント ベースのアドインのアクティブ化が有効になります。 詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Outlook**: イベント ベースのアクティブ化は現在 [プレビュー中で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 、Outlook on the web でのみ使用できます。 詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 使用できる文字数は 32 文字以内で、要素内の要素の属性と一 `resid` `id` `Url` 致する必要 `Resources` があります。 |
|  **lifetime**  |  いいえ  | 既定値は次 `lifetime` の `short` 値で、指定する必要があります。 Outlook アドインは値のみを使用 `short` します。 Excel アドインで共有ランタイムを使用する場合は、値を明示的に設定します `long` 。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
