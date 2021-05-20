---
title: マニフェスト ファイル内の起動イベント (プレビュー)
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するようにアドインを構成します。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555312"
---
# <a name="launchevent-element-preview"></a>起動イベント要素 (プレビュー)

サポートされているイベントに基づいてアクティブ化するようにアドインを構成します。 要素の子 [`<LaunchEvents>`](launchevents.md) 。 詳細については、「イベント[ベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。

**アドインの種類:** メール

> [!IMPORTANT]
> イベントベースのアクティブ化は現在[プレビュー段階にあり](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)、web 上およびWindowsでOutlookでのみ使用できます。 詳細については、「 [イベント ベースのアクティブ化機能をプレビューする方法](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)」を参照してください。

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
|  **Type**  |  はい  | サポートされているイベントの種類を指定します。 サポートされている種類のセットについては、「 [イベントベースのアクティブ化機能をプレビューする方法](../../outlook/autolaunch.md#supported-events)」を参照してください。 |
|  **FunctionName**  |  はい  | 属性で指定されたイベントを処理する JavaScript 関数の名前を指定します `Type` 。 |

## <a name="see-also"></a>関連項目

- [LaunchEvents](launchevents.md)
