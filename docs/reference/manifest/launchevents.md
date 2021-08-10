---
title: マニフェスト ファイルの LaunchEvents
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: c6714c4f52bdc1ed9d7a75a42100df8d3fe046c504575295880ff614fe4a447f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089749"
---
# <a name="launchevents-element"></a>LaunchEvents 要素

サポートされているイベントに基づいてアクティブ化するアドインを構成します。 要素の [`<ExtensionPoint>`](extensionpoint.md) 子。 詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** メール

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
| [LaunchEvent](launchevent.md) | はい |  サポートされているイベントを JavaScript ファイル内の関数にマップして、アドインのアクティブ化を行います。 |

## <a name="see-also"></a>関連項目

- [LaunchEvent](launchevent.md)
