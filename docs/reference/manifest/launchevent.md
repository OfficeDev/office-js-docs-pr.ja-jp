---
title: マニフェスト ファイルの LaunchEvent
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 23615424e194917a15b20ea4afbf7d9c5b8017e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152878"
---
# <a name="launchevent-element"></a>LaunchEvent 要素

サポートされているイベントに基づいてアクティブ化するアドインを構成します。 要素の [`<LaunchEvents>`](launchevents.md) 子。 詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](../../outlook/autolaunch.md)。

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **Type**  |  はい  | サポートされているイベントの種類を指定します。 サポートされている一連の種類については、「イベント ベースのライセンス認証Outlookアドインを構成する[」を参照してください](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  はい  | 属性で指定されたイベントを処理する JavaScript 関数の名前を指定 `Type` します。 |

## <a name="see-also"></a>関連項目

- [LaunchEvents](launchevents.md)
