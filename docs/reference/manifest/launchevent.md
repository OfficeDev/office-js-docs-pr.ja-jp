---
title: マニフェスト ファイルの LaunchEvent
description: LaunchEvent 要素は、サポートされているイベントに基づいてアクティブ化するアドインを構成します。
ms.date: 02/02/2022
ms.localizationpriority: medium
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
|  **種類**  |  はい  | サポートされているイベントの種類を指定します。 サポートされている一連の種類については、「イベント ベースのライセンス認証Outlookアドインを構成する[」を参照してください](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  はい  | 属性で指定されたイベントを処理する JavaScript 関数の名前を指定 `Type` します。 |
|  **SendMode** (プレビュー) |  不要  | 必須と `OnMessageSend` イベント `OnAppointmentSend` 。 アドインがアイテムの送信を停止する場合にユーザーが使用できるオプションを指定します。 使用可能なオプションについては、「使用可能な [SendMode オプション」を参照してください](#available-sendmode-options-preview)。 |

## <a name="available-sendmode-options-preview"></a>使用可能な SendMode オプション (プレビュー)

マニフェストにイベントを `OnMessageSend` 含 `OnAppointmentSend` める場合は、 **SendMode プロパティも設定する必要** があります。 使用可能なオプションを次に示します。 アドインが探している条件に基づいて、ユーザーは、アドインが送信されるアイテムに問題を見つけた場合に警告を受け取る。

| SendMode オプション | 説明 |
|---|---|
|`PromptUser`|アラートで、ユーザーは [任意の方法で送信] を選択するか、問題に対処してから、アイテムの再送信を試みます。|
|`SoftBlock`|ユーザーは、アイテムを再送信する前に問題を解決する必要があります。|

## <a name="see-also"></a>関連項目

- [LaunchEvents](launchevents.md)
- [イベント ベースのOutlook用にアドインを構成する](../../outlook/autolaunch.md#supported-events)
- [スマート アラートと OnMessageSend イベントをアドインOutlook使用する](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
