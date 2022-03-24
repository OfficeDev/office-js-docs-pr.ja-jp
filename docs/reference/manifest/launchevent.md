---
title: マニフェスト ファイルの LaunchEvent
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 03/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 71469693bff7213455582a3247778cabf92c2aa3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745814"
---
# <a name="launchevent-element"></a>LaunchEvent 要素

サポートされているイベントに基づいてアクティブ化するアドインを構成します。 要素の子 [`<LaunchEvents>`](launchevents.md) 。 詳細については、「イベント ベース[のアクティブ化Outlookアドインを構成する」を参照してください](../../outlook/autolaunch.md)。

**アドインの種類:** メール

**次の VersionOverrides スキーマでのみ有効です**。

- メール 1.1

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

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
|  **種類**  |  はい  | サポートされているイベントの種類を指定します。 サポートされる一連の種類については、「イベント ベースのライセンス認証Outlookアドインを構成する[」を参照してください](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  はい  | 属性で指定されたイベントを処理する JavaScript 関数の名前を指定 `Type` します。 |
|  **SendMode** (プレビュー) |  不要  | イベントによって使用`OnMessageSend``OnAppointmentSend`されます。 アドインがアイテムの送信を停止した場合、またはアドインが使用できない場合に、ユーザーが使用できるオプションを指定します。 **SendMode プロパティが** 含まれていない場合、オプションは`SoftBlock`既定で設定されます。 使用可能なオプションについては、「使用可能な [SendMode オプション」を参照してください](#available-sendmode-options-preview)。 |

## <a name="available-sendmode-options-preview"></a>使用可能な SendMode オプション (プレビュー)

マニフェストにイベントを `OnMessageSend` 含 `OnAppointmentSend` める場合は、 **SendMode プロパティも設定する必要** があります。 **SendMode プロパティが** 含まれていない場合、オプションは`SoftBlock`既定で設定されます。 使用可能なオプションを次に示します。 アドインが探している条件に基づいて、ユーザーは、アドインが送信されるアイテムに問題を見つけた場合に警告を受け取る。

| SendMode オプション | 説明 |
|---|---|
|`PromptUser`|アイテムがアドインの条件を満たしない場合、ユーザーはアラートで [任意に送信] を選択するか、問題に対処してから、アイテムの再送信を試みます。 アドインがアイテムの処理に時間がかかっている場合は、アドインの実行を停止し、[任意に送信] を選択するオプションが表示 **されます。** アドインが使用できない場合 (たとえば、アドインの読み込み中にエラーが発生した場合)、アイテムが送信されます。|
|`SoftBlock`|SendMode プロパティが **含まれていない** 場合の既定のオプション。 ユーザーは、送信するアイテムがアドインの条件を満たし、アイテムを再送信する前に問題に対処する必要があるという警告を受け取ります。 ただし、アドインが使用できない場合 (たとえば、アドインの読み込み中にエラーが発生した場合)、アイテムが送信されます。|
|`Block`|次の状況が発生した場合、アイテムは送信されません。<br>- アイテムがアドインの条件を満たしています。<br>- アドインはサーバーに接続できません。<br>- アドインの読み込み中にエラーが発生しました。|

## <a name="see-also"></a>関連項目

- [LaunchEvents](launchevents.md)
- [イベント ベースのOutlookアドインを構成する](../../outlook/autolaunch.md#supported-events)
- [スマート アラートと OnMessageSend イベントをアドインOutlook使用する](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
