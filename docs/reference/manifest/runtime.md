---
title: マニフェストファイル内のランタイム
description: Runtime 要素は、アドインが、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに対して共有 JavaScript ランタイムを使用するように構成します。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159368"
---
# <a name="runtime-element-preview"></a>Runtime 要素 (プレビュー)

共有された JavaScript ランタイムを使用するようにアドインを構成し、さまざまなコンポーネントがすべて同じランタイムで実行されるようにします。 要素の子 [`<Runtimes>`](runtimes.md) 。

Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。 詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Outlook**: イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。 詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。

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
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 は、 `resid` `id` 要素内の要素の属性と一致している必要があり `Url` `Resources` ます。 |
|  **時間**  |  不要  | の既定値は、を `lifetime` `short` 指定する必要はありません。 Outlook アドインは、値のみを使用し `short` ます。 Excel アドインで共有ランタイムを使用する場合は、の値をに明示的に設定し `long` ます。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
