---
title: マニフェストファイル内の LaunchEvents (プレビュー)
description: LaunchEvents 要素は、サポートされているイベントに基づいてアクティブになるようにアドインを構成します。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278556"
---
# <a name="launchevents-element-preview"></a>LaunchEvents 要素 (プレビュー)

サポートされているイベントに基づいて、アドインをアクティブにするように構成します。 要素の子 [`<ExtensionPoint>`](extensionpoint.md) 。 詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

**アドインの種類:** メール

> [!IMPORTANT]
> イベントベースのライセンス認証は現在[プレビュー段階で](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)あり、web 上の Outlook でのみ使用できます。 詳細については、「[イベントベースのライセンス認証機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。

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

[Extensionpoint](extensionpoint.md) (**launchevent**メールアドイン)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | はい |  アドインをアクティブ化するために、JavaScript ファイルの関数にサポートされているイベントをマップします。 |

## <a name="see-also"></a>関連項目

- [LaunchEvent](launchevent.md)
