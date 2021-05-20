---
title: マニフェスト ファイルのランタイム
description: Runtime 要素は、リボン、作業ウィンドウ、カスタム関数など、さまざまなコンポーネントに共有 JavaScript ランタイムを使用するようにアドインを構成します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555305"
---
# <a name="runtime-element"></a>ランタイム要素

共有 JavaScript ランタイムを使用して、さまざまなコンポーネントがすべて同じランタイムで実行されるようにアドインを構成します。 要素の子 [`<Runtimes>`](runtimes.md) 。

**アドインの種類:** 作業ウィンドウ,メール

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
| [上書き](override.md) (プレビュー) | いいえ | **Outlook**:[デスクトップで起動イベント拡張ポイント](../../reference/manifest/extensionpoint.md#launchevent-preview)ハンドラーに必要な JavaScript ファイルの URL の場所 Outlookを指定します。 **重要**: 現在、定義できる要素は 1 つのみ `<Override>` で、型が必要です `javascript` 。|

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid**  |  はい  | アドインの HTML ページの URL の場所を指定します。 は `resid` 32 文字以内 `id` で、要素の属性と一致する必要があります `Url` `Resources` 。 |
|  **一生**  |  いいえ  | デフォルト値 `lifetime` は `short` 、指定する必要はありません。 Outlookアドインでは、値のみを使用します `short` 。 Excel アドインで共有ランタイムを使用する場合は、値を明示的に に 設定 `long` します。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtimes.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)
