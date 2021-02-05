---
title: マニフェスト ファイル内の LaunchEvents (プレビュー)
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 9df059879018d79a61f1c900888c8d197e0b9880
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104813"
---
# <a name="launchevents-element-preview"></a>LaunchEvents 要素 (プレビュー)

サポートされているイベントに基づいてアクティブ化するアドインを構成します。 要素の [`<ExtensionPoint>`](extensionpoint.md) 子。 詳細については、「イベント ベースのアクティブ [化用に Outlook アドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** メール

> [!IMPORTANT]
> イベント ベースのアクティブ化は現在 [プレビュー中であり](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 、Outlook on the web および Windows でのみ使用できます。 詳細については、イベント ベースの [アクティブ化機能をプレビューする方法を参照してください](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

## <a name="syntax"></a>構文

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a>含まれる場所

[ExtensionPoint](extensionpoint.md) (**LaunchEvent** メール アドイン)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | はい |  アドインのアクティブ化のために、サポートされているイベントを JavaScript ファイル内の関数にマップします。 |

## <a name="see-also"></a>関連項目

- [LaunchEvent](launchevent.md)
