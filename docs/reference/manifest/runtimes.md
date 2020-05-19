---
title: マニフェストファイル内のランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 22156a171ca2f423024efb1b3d2a6fdae07dfef6
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278365"
---
# <a name="runtimes-element"></a>ランタイム要素

アドインの実行時のランタイムを指定します。 要素の子 [`<Host>`](host.md) 。

Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。 詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Excel**: 共有ランタイムは現在プレビュー段階であり、Windows 上の Excel でのみ使用できます。 プレビュー機能を試すには、 [Office Insider](https://insider.office.com/)に参加する必要があります。
>
> **Outlook**: イベントベースのライセンス認証機能は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。 詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。

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
