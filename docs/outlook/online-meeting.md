---
title: オンライン会議プロバイダー Outlookモバイル アドインを作成する
description: オンライン会議サービス プロバイダー Outlookモバイル アドインをセットアップする方法について説明します。
ms.topic: article
ms.date: 07/09/2021
localization_priority: Normal
ms.openlocfilehash: 34574809e2b874217113e42043b3bde7ef0dd8ba
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936340"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>オンライン会議プロバイダー Outlookモバイル アドインを作成する

オンライン会議のセットアップは、Outlook ユーザーの主要なエクスペリエンスであり、モバイルユーザーとのTeams作成[Outlookです。](/microsoftteams/teams-add-in-for-outlook) ただし、Microsoft 以外のサービスを使用Outlookオンライン会議を作成すると、面倒な場合があります。 この機能を実装することで、サービス プロバイダーは、アドイン ユーザーのオンライン会議Outlookを合理化できます。

> [!IMPORTANT]
> この機能は、Android と iOS でのみサポートされ、サブスクリプションMicrosoft 365されます。

この記事では、オンライン会議サービスを使用してユーザーが会議を整理して参加できるよう、Outlook モバイル アドインをセットアップする方法について学習します。 この記事では、架空のオンライン会議サービス プロバイダー "Contoso" を使用します。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

ユーザーがアドインを使用してオンライン会議を作成するには、親要素の下のマニフェストで [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) 拡張ポイントを構成する必要があります `MobileFormFactor` 。 他のフォーム ファクターはサポートされていません。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にあるファイルを開きます。

1. ノード全体 (開 `<VersionOverrides>` くタグと閉じるタグを含む) を選択し、次の XML に置き換えてください。

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> Outlook アドインのマニフェストの詳細については[、「Outlook](manifests.md)アドイン マニフェスト」および「Outlook Mobile 用アドイン コマンドのサポート[の追加」を参照してください](add-mobile-support.md)。

## <a name="implement-adding-online-meeting-details"></a>オンライン会議の詳細の追加を実装する

このセクションでは、アドイン スクリプトでユーザーの会議を更新して、オンライン会議の詳細を含める方法について説明します。

1. 同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。

1. ファイルのコンテンツ全体を **次commands.js** JavaScript に置き換える。

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a>テストと検証

アドインをテストして検証 [するには、通常のガイダンスに従います](testing-and-tips.md)。 [Android、Outlook on the web、Windows](sideload-outlook-add-ins-for-testing.md) Mac でサイドローディングした後、Android Outlook iOS モバイル デバイスでデバイスを再起動します。 次に、新しい会議画面で、Microsoft TeamsまたはSkypeが自分のトグルに置き換えられるか確認します。

### <a name="create-meeting-ui"></a>会議 UI の作成

会議の開催者として、会議を作成すると、次の 3 つの画像のような画面が表示されます。

[![Android の会議の作成画面 - Contoso のトグル オフ。](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Android の会議の作成画面 - Contoso の読み込みトグル。](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Android の会議の作成画面 - Contoso のトグル オン。](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>会議の UI に参加する

会議の出席者として、会議を表示すると、次のような画面が表示されます。

[![Android の参加会議画面のスクリーンショット。](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> [参加] リンクが表示されない場合は、サービスのオンライン会議テンプレートがサーバーに登録されていない可能性があります。 詳細については [、「オンライン会議テンプレートの登録」](#register-your-online-meeting-template) セクションを参照してください。

## <a name="register-your-online-meeting-template"></a>オンライン会議テンプレートを登録する

サービスのオンライン会議テンプレートを登録する場合は、詳細に関する問題GitHub作成できます。 その後、登録のタイムラインを調整するためにお問い合わせください。

1. この記事の **最後** にある [フィードバック] セクションに移動します。
1. [このページ **] リンクを押** します。
1. 新しい **問題のタイトル** を "my-service のオンライン会議テンプレートを登録する" に設定し、サービス名 `my-service` に置き換える。
1. 問題本文で、文字列 "[Enter feedback here]" を、この記事の「オンライン会議の詳細の追加を実装する」セクションの類似の変数で設定した文字列に `newBody` 置き換える必要があります。 [](#implement-adding-online-meeting-details)
1. [新 **しい問題の送信] をクリックします**。

![Contoso のサンプル コンテンツGitHub新しい問題の画面のスクリーンショット。](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>使用可能な API

この機能では、次の API を使用できます。

- 予定オーガナイザー API
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getAsync_coercionType__options__callback_), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setAsync_data__options__callback_))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalAttendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredAttendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- 認証フローの処理
  - [ダイアログ API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>制限

いくつかの制限が適用されます。

- オンライン会議サービス プロバイダーにのみ適用されます。
- 管理者がインストールしたアドインだけが会議の作成画面に表示され、既定の構成オプションまたは TeamsオプションSkypeされます。 ユーザーがインストールしたアドインはアクティブ化されません。
- アドイン アイコンは、16 進数コードまたは他の色形式で同等の値を使用してグレー `#919191` [スケールで表示する必要があります](https://convertingcolors.com/hex-color-919191.html)。
- 予定オーガナイザー (作成) モードでは、1 つの UI レス コマンドだけがサポートされます。
- アドインは、1 分のタイムアウト期間内に予定フォームの会議の詳細を更新する必要があります。 ただし、認証用に開いたアドインなどのダイアログ ボックスで費やされた時間は、タイムアウト期間から除外されます。

## <a name="see-also"></a>関連項目

- [Outlook Mobile のアドイン](outlook-mobile-addins.md)
- [Outlook Mobile のアドイン コマンドのサポートを追加する](add-mobile-support.md)
