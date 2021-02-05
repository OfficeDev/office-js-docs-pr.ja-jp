---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 74bb2b432f46d5876601052003e20ff843e13b06
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104827"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の [`<Host>`](host.md) 子。

> [!NOTE]
> Windows 上の Officeで実行する場合、アドインは Internet Explorer 11 ブラウザーを使用します。

Excel では、この要素により、リボン、作業ウィンドウ、およびカスタム関数で同じランタイムを使用できます。 詳細については、「共有 [JavaScript ランタイムを使用するために Excel](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する」を参照してください。

Outlook では、この要素により、イベント ベースのアドインのアクティブ化が有効になります。 詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Outlook**: イベント ベースのアクティブ化 [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)機能は現在プレビュー中で、Outlook on the web および Windows でのみ使用できます。 詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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
| [Runtime](runtime.md) | はい |  アドインのランタイム。 |

## <a name="see-also"></a>関連項目

- [Runtime](runtime.md)
