---
title: Outlook モバイル アドインで外部アプリケーションに予定ノートを記録する
description: 予定メモやその他の詳細を外部アプリケーションにログに記録する Outlook モバイル アドインを設定する方法について説明します。
ms.topic: article
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: a980b68c603154c42112f525ec6285b740ce38a5
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607584"
---
# <a name="log-appointment-notes-to-an-external-application-in-outlook-mobile-add-ins"></a>Outlook モバイル アドインで外部アプリケーションに予定ノートを記録する

予定メモやその他の詳細を顧客関係管理 (CRM) またはメモ作成アプリケーションに保存すると、出席した会議を追跡するのに役立ちます。

この記事では、Outlook モバイル アドインを設定して、ユーザーが自分の予定に関するメモやその他の詳細を CRM またはメモ作成アプリケーションに記録できるようにする方法について説明します。 この記事では、"Contoso" という架空の CRM サービス プロバイダーを使用します。

> [!IMPORTANT]
> この機能は、Microsoft 365 サブスクリプションを使用する Android でのみサポートされます。

## <a name="set-up-your-environment"></a>環境を設定する

Office アドイン用 Yeoman ジェネレーターを使用してアドイン プロジェクトを作成するには [、Outlook クイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) を完了します。

## <a name="capture-and-view-appointment-notes"></a>予定のメモをキャプチャして表示する

関数コマンドまたは作業ウィンドウを実装することを選択できます。 アドインを更新するには、関数コマンドまたは作業ウィンドウのタブを選択し、指示に従います。

# <a name="function-command"></a>[関数コマンド](#tab/noui)

このオプションを使用すると、リボンから関数コマンドを選択したときに、ユーザーは自分の予定に関するメモやその他の詳細をログに記録して表示できます。

### <a name="configure-the-manifest"></a>マニフェストを構成する

ユーザーがアドインで予定ノートをログに記録できるようにするには、親要素`MobileFormFactor`の下のマニフェストで [MobileLogEventAppointmentAttendee 拡張ポイント](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)を構成する必要があります。 その他のフォーム ファクターはサポートされていません。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. ノード全体 `<VersionOverrides>` (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えます。 **Contoso** への参照はすべて、会社の情報に置き換えてください。

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
              <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="apptReadGroup">
                    <Label resid="residDescription"/>
                    <Control xsi:type="Button" id="apptReadOpenPaneButton">
                      <Label resid="residLabel"/>
                      <Supertip>
                        <Title resid="residLabel"/>
                        <Description resid="residTooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="icon-16"/>
                        <bt:Image size="32" resid="icon-32"/>
                        <bt:Image size="80" resid="icon-80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>logCRMEvent</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
            <MobileFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <FunctionName>logCRMEvent</FunctionName>
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
            <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
            <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Outlook アドインのマニフェストの詳細については、「Outlook アドイン マニフェスト」と [「Outlook](manifests.md) [Mobile 用アドイン コマンドのサポートの追加](add-mobile-support.md)」を参照してください。

### <a name="capture-appointment-notes"></a>予定のメモをキャプチャする

このセクションでは、ユーザーが **[ログ** ] ボタンを選択したときにアドインが予定の詳細を抽出する方法について説明します。

1. 同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/commands/commands.js** を開きます。

1. **commands.js** ファイルの内容全体を次の JavaScript に置き換えます。

    ```js
    var event;

    Office.initialize = function (reason) {
      // Add any initialization code here.
    };

    function logCRMEvent(appointmentEvent) {
      event = appointmentEvent;
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
          } else {
            console.error("Failed to get body.");
            event.completed({ allowEvent: false });
          }
        }
      );
    }

    // Register the function.
    Office.actions.associate("logCRMEvent", logCRMEvent);
    ```

次に、 **commands.html** ファイルを参照 **commands.js** 更新します。

1. 同じクイック スタート プロジェクトから、コード エディターで **./src/commands/commands.html** ファイルを開きます。

1. 検索して、次のように置き換えます `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` 。

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="commands.js"></script>
    ```

### <a name="view-appointment-notes"></a>予定のメモを表示する

**ログ** ボタンラベルを切り替えて **表示** するには、この目的のために予約されている **EventLogged** カスタム プロパティを設定します。 ユーザーが **[表示** ] ボタンを選択すると、この予定のログに記録されたメモを確認できます。

アドインは、ログ表示エクスペリエンスを定義します。 たとえば、ユーザーが **[表示** ] ボタンを選択したときに、ログに記録された予定ノートをダイアログに表示できます。 ダイアログの使用の詳細については、「 [Office アドインで Office ダイアログ API を使用する」](../develop/dialog-api-in-office-add-ins.md)を参照してください。

次の関数を **./src/commands/commands.js** に追加します。 この関数は、現在の予定アイテムに **EventLogged** カスタム プロパティを設定します。

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
              event.completed({ allowEvent: true });
              event = undefined;
            }
          }
        );
      }
    }
  );
}
```

次に、アドインが予定ノートを正常にログに記録した後、それを呼び出します。 たとえば、次の関数に示すように **、logCRMEvent** から呼び出すことができます。

```js
function logCRMEvent(appointmentEvent) {
  event = appointmentEvent;
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Replace `event.completed({ allowEvent: true });` with the following statement.
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
        event.completed({ allowEvent: false });
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>予定ログを削除する

ユーザーがログを元に戻すか、ログに記録された予定のメモを削除して置換ログを保存できるようにする場合は、2 つのオプションがあります。

1. ユーザーがリボンの適切なボタンを選択したときに、Microsoft Graph を使用して [カスタム プロパティ オブジェクトをクリア](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) します。
1. 次の関数を **./src/commands/commands.js** に追加して、現在の予定アイテムの **EventLogged** カスタム プロパティをクリアします。

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                  event.completed({ allowEvent: true });
                  event = undefined;
                }
              }
            );
          }
        }
      );
    }
    ```

次に、カスタム プロパティをクリアするときに呼び出します。 たとえば、次の関数に示すように、ログの設定が何らかの方法で失敗した場合は、 **logCRMEvent** から呼び出すことができます。

  ```js
  function logCRMEvent(appointmentEvent) {
    event = appointmentEvent;
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          // Replace `event.completed({ allowEvent: false });` with the following statement.
          clearCustomProperties();
        }
      }
    );
  }
  ```

# <a name="task-pane"></a>[作業ウィンドウ](#tab/taskpane)

このオプションを使用すると、ユーザーは作業ウィンドウから自分の予定に関するメモやその他の詳細をログに記録して表示できます。

### <a name="configure-the-manifest"></a>マニフェストを構成する

ユーザーがアドインで予定ノートをログに記録できるようにするには、親要素`MobileFormFactor`の下のマニフェストで [MobileLogEventAppointmentAttendee 拡張ポイント](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)を構成する必要があります。 その他のフォーム ファクターはサポートされていません。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. ノード全体 `<VersionOverrides>` (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えます。 **Contoso** への参照はすべて、会社の情報に置き換えてください。

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
                <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                  <OfficeTab id="TabDefault">
                    <Group id="apptReadGroup">
                      <Label resid="residDescription"/>
                      <Control xsi:type="Button" id="apptReadOpenPaneButton">
                        <Label resid="residLabel"/>
                        <Supertip>
                          <Title resid="residLabel"/>
                          <Description resid="residTooltip"/>
                        </Supertip>
                        <Icon>
                          <bt:Image size="16" resid="icon-16"/>
                          <bt:Image size="32" resid="icon-32"/>
                          <bt:Image size="80" resid="icon-80"/>
                        </Icon>
                        <Action xsi:type="ShowTaskpane">
                          <SourceLocation resid="Taskpane.Url"/>
                        </Action>
                      </Control>
                    </Group>
                  </OfficeTab>
                </ExtensionPoint>
              </DesktopFormFactor>
              <MobileFormFactor>
                <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
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
              <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
              <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
              <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
              <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
              <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
            </bt:ShortStrings>
            <bt:LongStrings>
              <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
            </bt:LongStrings>
          </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Outlook アドインのマニフェストの詳細については、「Outlook アドイン マニフェスト」と [「Outlook](manifests.md) [Mobile 用アドイン コマンドのサポートの追加](add-mobile-support.md)」を参照してください。

### <a name="capture-appointment-notes"></a>予定のメモをキャプチャする

このセクションでは、ユーザーが [ **ログ** ] ボタンを選択したときに、ログに記録された予定のメモとその他の詳細を作業ウィンドウに表示する方法について説明します。

1. 同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/taskpane/taskpane.js** を開きます。

1. **taskpane.js** ファイルの内容全体を次の JavaScript に置き換えます。

    ```js
    // Office is ready.
    Office.onReady(function () {
        getEventData();
      }
    );

    function getEventData() {
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("event logged successfully");
          } else {
            console.error("Failed to get body.");
          }
        }
      );
    }
    ```

次に、 **taskpane.html** ファイルを参照 **taskpane.js** 更新します。

1. 同じクイック スタート プロジェクトから、コード エディターで **./src/taskpane/taskpane.html** ファイルを開きます。

1. 検索して、次のように置き換えます `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` 。

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="taskpane.js"></script>
    ```

### <a name="view-appointment-notes"></a>予定のメモを表示する

**ログ** ボタンラベルを切り替えて **表示** するには、この目的のために予約されている **EventLogged** カスタム プロパティを設定します。 ユーザーが **[表示** ] ボタンを選択すると、この予定のログに記録されたメモを確認できます。 アドインは、ログ表示エクスペリエンスを定義します。

次の関数を **./src/taskpane/taskpane.js** に追加します。 この関数は、現在の予定アイテムに **EventLogged** カスタム プロパティを設定します。

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
            }
          }
        );
      }
    }
  );
}
```

次に、アドインが予定ノートを正常にログに記録した後、それを呼び出します。 たとえば、次の関数に示すように **、getEventData** から呼び出すことができます。

```js
function getEventData() {
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("event logged successfully");
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>予定ログを削除する

ユーザーがログを元に戻すか、ログに記録された予定のメモを削除して置換ログを保存できるようにする場合は、2 つのオプションがあります。

1. ユーザーが作業ウィンドウで適切なボタンを選択したときに、Microsoft Graph を使用して [カスタム プロパティ オブジェクトをクリア](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) します。
1. 次の関数を **./src/taskpane/taskpane.js** に追加して、現在の予定アイテムの **EventLogged** カスタム プロパティをクリアします。

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                }
              }
            );
          }
        }
      );
    }
    ```

次に、カスタム プロパティをクリアするときに呼び出します。 たとえば、次の関数に示すように、ログの設定に何らかの方法で失敗した場合は **、getEventData** から呼び出すことができます。

  ```js
  function getEventData() {
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("event logged successfully");
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          clearCustomProperties();
        }
      }
    );
  }
  ```

---

## <a name="test-and-validate"></a>テストと検証

1. アドインを [テストして検証するには、通常のガイダンスに](testing-and-tips.md)従います。
1. アドインを Outlook on the web、Windows、または Mac で[サイドロード](sideload-outlook-add-ins-for-testing.md)した後、Android モバイル デバイスで Outlook を再起動します。
1. 出席者として予定を開き、 **Meeting Insights** カードの下に[ **ログ** ] ボタンと共にアドインの名前を含む新しいカードがあることを確認します。

### <a name="ui-log-the-appointment-notes"></a>UI: 予定のメモを記録する

会議出席者は、会議を開くときに次の図のような画面が表示されます。

![Android の予定画面の [ログ] ボタンを示すスクリーンショット。](../images/outlook-android-log-appointment-details.jpg)

### <a name="ui-view-the-appointment-log"></a>UI: 予定ログを表示する

予定ノートを正常にログに記録した後、ボタンに **[ログ**] ではなく [**表示]** というラベルが付けられます。 次の図のような画面が表示されます。

![Android の予定画面の [表示] ボタンを示すスクリーンショット。](../images/outlook-android-view-appointment-log.jpg)

## <a name="available-apis"></a>使用可能な API

この機能では、次の API を使用できます。

- [ダイアログ API](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- 次 **を除く**[予定の読み取り (出席者) API](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true):
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## <a name="restrictions"></a>制限

いくつかの制限が適用されます。

- **[ログ]** ボタンの名前は変更できません。 ただし、予定アイテムにカスタム プロパティを設定することで、別のラベルを表示する方法があります。 詳細については、必要に応じて [、関数コマンド](?tabs=noui#view-appointment-notes)または [作業ウィンドウ](?tabs=taskpane#view-appointment-notes-1)の **予定ノートの表示** セクションを参照してください。
- **[ログ**] ボタンのラベルを **[表示** と戻り] に切り替える場合は、**EventLogged** カスタム プロパティを使用する必要があります。
- アドイン アイコンは、16 進コード `#919191` を使用するか、 [他の色形式](https://convertingcolors.com/hex-color-919191.html)で同等の色を使用してグレースケールにする必要があります。
- アドインは、1 分間のタイムアウト期間内に予定フォームから会議の詳細を抽出する必要があります。 ただし、認証用に開かれたアドインがダイアログ ボックスで費やされた時間は、タイムアウト期間から除外されます。

## <a name="see-also"></a>関連項目

- [Outlook Mobile のアドイン](outlook-mobile-addins.md)
- [Outlook Mobile のアドイン コマンドのサポートを追加する](add-mobile-support.md)
