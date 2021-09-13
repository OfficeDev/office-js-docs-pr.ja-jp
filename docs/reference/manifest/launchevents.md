---
title: マニフェスト ファイルの LaunchEvents
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 05/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 02e0b21d65733492a783ffb099caf9e76225e53f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151248"
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
