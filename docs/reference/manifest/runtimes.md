---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789164"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の [`<Host>`](host.md) 子。

> [!NOTE]
> Windows Officeで実行する場合、アドインは Internet Explorer 11 ブラウザーを使用します。

Excel では、この要素により、リボン、作業ウィンドウ、およびカスタム関数で同じランタイムを使用できます。 詳細については、「共有 [JavaScript ランタイムを使用するために Excel](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する」を参照してください。

Outlook では、この要素によってイベント ベースのアドインのアクティブ化が有効になります。 詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Outlook**: イベント ベースのアクティブ化 [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)機能は現在プレビュー中で、Outlook on the web でのみ使用できます。 詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

[Host](host.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [ランタイム](runtime.md) | はい |  アドインのランタイム。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
