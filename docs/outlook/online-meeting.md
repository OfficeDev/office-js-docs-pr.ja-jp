---
title: オンライン会議プロバイダー用の Outlook モバイルアドインを作成する (プレビュー)
description: オンライン会議サービスプロバイダー用の Outlook mobile アドインをセットアップする方法について説明します。
ms.topic: article
ms.date: 04/21/2020
localization_priority: Normal
ms.openlocfilehash: 5fd0b28a661f6d2e8f3084427920c1a31053ae5b
ms.sourcegitcommit: 3355c6bd64ecb45cea4c0d319053397f11bc9834
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/22/2020
ms.locfileid: "43744860"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider-preview"></a>オンライン会議プロバイダー用の Outlook モバイルアドインを作成する (プレビュー)

オンライン会議の設定は、Outlook ユーザーにとって中心的な操作であり、Outlook mobile[を使用して Teams 会議を](/microsoftteams/teams-add-in-for-outlook)簡単に作成できます。 ただし、Microsoft 以外のサービスを使用して Outlook でオンライン会議を作成するのは煩雑な場合があります。 この機能を実装することにより、サービスプロバイダーは、Outlook アドインユーザーに対してオンライン会議の作成環境を合理化することができます。

> [!NOTE]
> この機能は、Office 365 サブスクリプションを使用した Android の[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。

この記事では、ユーザーがオンライン会議サービスを使用して会議を整理し、会議に参加できるようにするために Outlook モバイルアドインをセットアップする方法について説明します。 この記事全体で、架空のオンライン会議サービスプロバイダーである "Contoso" を使用します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

ユーザーがアドインを使用してオンライン会議を作成できるようにするには`MobileOnlineMeetingCommandSurface` 、マニフェストで親要素`MobileFormFactor`の下に拡張点を構成する必要があります。 その他のフォームファクターはサポートされていません。

次の例は、 `MobileFormFactor`要素と`MobileOnlineMeetingCommandSurface`拡張点を含むマニフェストからの抜粋を示しています。

> [!TIP]
> Outlook アドインのマニフェストの詳細については、「outlook[アドインのマニフェスト](manifests.md)」および「 [outlook Mobile のアドインコマンドのサポートを追加](add-mobile-support.md)する」を参照してください。

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a>オンライン会議の詳細の追加を実装する

このセクションでは、アドインスクリプトでユーザーの会議を更新してオンライン会議の詳細を含める方法について説明します。

次の例は、オンライン会議の詳細を作成する方法を示しています。 ここでは、サービスから会議開催者の ID とその他の詳細を取得する方法については説明しません。

```js
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
```

次の例は、マニフェストで`insertContosoMeeting`参照される UI を使用しない関数を定義して、オンライン会議の詳細で会議の本文を更新する方法を示しています。

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

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
```

次の例は、前の例で使用`updateBody`した、会議の現在の本文にオンライン会議の詳細を追加する、サポート関数の実装を示しています。

```js
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
```

## <a name="testing-and-validation"></a>テストと検証

通常のガイダンスに従って、[アドインをテストし、検証](testing-and-tips.md)します。 Outlook on the web、Windows、または Mac で[サイド](sideload-outlook-add-ins-for-testing.md)ロード後に、android モバイルデバイスで outlook を再起動します (現時点でサポートされている唯一のクライアントは android です)。 次に、新しい会議画面で、Microsoft Teams または Skype のトグルが自分のものに置き換えられていることを確認します。

### <a name="create-meeting-ui"></a>会議 UI を作成する

会議の開催者として、会議を作成するときに次の3つの画像のような画面が表示されます。

android の[ ![](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [[会議を作成する] 画面のスクリーンショット-contoso-contoso の [会議を作成する] 画面のスクリーンショットを非表示にする-contoso の会議![](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![画面を作成する](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>ミーティング UI への参加

会議の出席者として、会議を表示するときに次のような画面が表示されます。

[![Android に参加した会議画面のスクリーンショット](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

## <a name="available-apis"></a>利用可能な Api

この機能では、次の Api を使用できます。

- 予定の開催者 Api
  - [Office. アイテムの件名](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject)([subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))
  - 。[開始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start)(時刻) ([時間](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - (時刻)-[終了](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end)([時間](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - [Office. メールボックス. アイテムの場所](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location)([場所](/javascript/api/outlook/office.location?view=outlook-js-preview))
  - [Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) -任意の出席者 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ....../の内容 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [GetAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-)(body, [Body, setasync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-)) (添付[コンテンツ](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body)の添付)
  - [CustomProperties (](/javascript/api/outlook/office.customproperties?view=outlook-js-preview)) のようにし[ます。](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-)
  - [RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))
- 認証フローを処理する
  - [ダイアログ API](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>制限

いくつかの制限が適用されます。

- オンライン会議サービスプロバイダーにのみ適用されます。
- 現在プレビュー中であるため、この機能は運用アドインでは使用しないでください。
- 現時点では、Android はサポートされている唯一のクライアントです。 IOS でのサポートは近日に予定されています。
- 既定の Teams または Skype オプションを置き換えて、管理者によってインストールされたアドインのみが会議の作成画面に表示されます。 ユーザーがインストールしたアドインはアクティブ化されません。
- アドインアイコンは、16進コード`#919191`または[その他の色の形式](https://convertingcolors.com/hex-color-919191.html)の同等機能を使用したグレースケールである必要があります。
- 予定の開催者 (新規作成) モードでは、UI レスコマンドは1つだけサポートされています。

## <a name="see-also"></a>関連項目

- [Outlook Mobile のアドイン](outlook-mobile-addins.md)
- [Outlook Mobile のアドイン コマンドのサポートを追加する](add-mobile-support.md)
