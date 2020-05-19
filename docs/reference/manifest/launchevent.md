---
title: マニフェストファイルの LaunchEvent (プレビュー)
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブになるようにアドインを構成します。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278559"
---
# <a name="launchevent-element-preview"></a>LaunchEvent 要素 (プレビュー)

サポートされているイベントに基づいて、アドインをアクティブにするように構成します。 要素の子 [`<LaunchEvents>`](launchevents.md) 。 詳細については、「[イベントベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **種類**  |  はい  | サポートされているイベントの種類を指定します。 使用できる型は `OnNewMessageCompose` 、および `OnNewAppointmentOrganizer` です。 |
|  **FunctionName**  |  はい  | 属性で指定されたイベントを処理する JavaScript 関数の名前を指定し `Type` ます。 |

## <a name="see-also"></a>関連項目

- [LaunchEvents](launchevents.md)
