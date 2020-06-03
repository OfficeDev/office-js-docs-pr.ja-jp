---
title: マニフェストファイル内のランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 95549d88df24a7d7c54cf27c92c15693491bdf29
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/03/2020
ms.locfileid: "44520340"
---
# <a name="runtimes-element"></a>ランタイム要素

アドインの実行時のランタイムを指定します。 要素の子 [`<Host>`](host.md) 。

> [!NOTE]
> Windows で Office を実行している場合、アドインは Internet Explorer 11 ブラウザーを使用します。

Excel では、この要素を使用すると、リボン、作業ウィンドウ、およびカスタム関数が同じランタイムを使用できるようになります。 詳細については、「[共有 JavaScript ランタイムを使用するように Excel アドインを構成する](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

Outlook では、この要素はイベントベースのアドインのアクティブ化を有効にします。 詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

**アドインの種類:** 作業ウィンドウ、メール

> [!IMPORTANT]
> **Excel**: 共有ランタイムは、現在 Windows 上の Excel でのみ使用できます。
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
