---
title: オンライン会議プロバイダー用の Outlook モバイルアドインを作成する
description: オンライン会議サービスプロバイダー用の Outlook mobile アドインをセットアップする方法について説明します。
ms.topic: article
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: d3dd1f035c69b668c05f80b36ef48108b8a9cecc
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431074"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="b9bea-103">オンライン会議プロバイダー用の Outlook モバイルアドインを作成する</span><span class="sxs-lookup"><span data-stu-id="b9bea-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="b9bea-104">オンライン会議の設定は、Outlook ユーザーにとって中心的な操作であり、Outlook mobile [を使用して Teams 会議を](/microsoftteams/teams-add-in-for-outlook) 簡単に作成できます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="b9bea-105">ただし、Microsoft 以外のサービスを使用して Outlook でオンライン会議を作成するのは煩雑な場合があります。</span><span class="sxs-lookup"><span data-stu-id="b9bea-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="b9bea-106">この機能を実装することにより、サービスプロバイダーは、Outlook アドインユーザーに対してオンライン会議の作成環境を合理化することができます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9bea-107">この機能は、Microsoft 365 サブスクリプションを使用した Android でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b9bea-107">This feature is only supported on Android with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="b9bea-108">この記事では、ユーザーがオンライン会議サービスを使用して会議を整理し、会議に参加できるようにするために Outlook モバイルアドインをセットアップする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="b9bea-109">この記事全体で、架空のオンライン会議サービスプロバイダーである "Contoso" を使用します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="b9bea-110">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="b9bea-110">Set up your environment</span></span>

<span data-ttu-id="b9bea-111">Outlook の [クイックスタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) に記入します。このアドインプロジェクトは、Office アドイン用の [アプリ] ジェネレーターを使用して作成されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b9bea-112">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="b9bea-112">Configure the manifest</span></span>

<span data-ttu-id="b9bea-113">ユーザーがアドインを使用してオンライン会議を作成できるようにするには、 `MobileOnlineMeetingCommandSurface` マニフェストで親要素の下に拡張点を構成する必要があり `MobileFormFactor` ます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-113">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="b9bea-114">その他のフォームファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="b9bea-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="b9bea-115">コードエディターで、[クイックスタート] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="b9bea-116">プロジェクトのルートにある **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="b9bea-117">`<VersionOverrides>`ノード全体 (open タグと close タグを含む) を選択し、次の XML に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="b9bea-118">Outlook アドインのマニフェストの詳細については、「outlook [アドインのマニフェスト](manifests.md) 」および「 [outlook Mobile のアドインコマンドのサポートを追加](add-mobile-support.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9bea-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="b9bea-119">オンライン会議の詳細の追加を実装する</span><span class="sxs-lookup"><span data-stu-id="b9bea-119">Implement adding online meeting details</span></span>

<span data-ttu-id="b9bea-120">このセクションでは、アドインスクリプトでユーザーの会議を更新してオンライン会議の詳細を含める方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="b9bea-121">同じクイックスタートプロジェクトから、コードエディターで **/src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="b9bea-122">**commands.js**ファイルの内容全体を次の JavaScript に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

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

## <a name="testing-and-validation"></a><span data-ttu-id="b9bea-123">テストと検証</span><span class="sxs-lookup"><span data-stu-id="b9bea-123">Testing and validation</span></span>

<span data-ttu-id="b9bea-124">通常のガイダンスに従って、 [アドインをテストし、検証](testing-and-tips.md)します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="b9bea-125">Outlook on the web、Windows、または Mac で [サイド](sideload-outlook-add-ins-for-testing.md) ロード後に、Android モバイルデバイスで outlook を再起動します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.</span></span> <span data-ttu-id="b9bea-126">(現時点でサポートされているクライアントは Android のみです)。次に、新しい会議画面で、Microsoft Teams または Skype のトグルが自分のものに置き換えられていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-126">(Android is the only supported client for now.) Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="b9bea-127">会議 UI を作成する</span><span class="sxs-lookup"><span data-stu-id="b9bea-127">Create meeting UI</span></span>

<span data-ttu-id="b9bea-128">会議の開催者として、会議を作成するときに次の3つの画像のような画面が表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="b9bea-129">android の[ ![ [会議を作成する] 画面](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox)のスクリーンショット-contoso-contoso の[ ![ [](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox)会議を作成する] 画面のスクリーンショットを非表示にする-contoso の会議[ ![ 画面を作成する](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="b9bea-129">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="b9bea-130">ミーティング UI への参加</span><span class="sxs-lookup"><span data-stu-id="b9bea-130">Join meeting UI</span></span>

<span data-ttu-id="b9bea-131">会議の出席者として、会議を表示するときに次のような画面が表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-131">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="b9bea-132">[![Android に参加した会議画面のスクリーンショット](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="b9bea-132">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9bea-133">[ **参加** ] リンクが表示されない場合は、サービスのオンライン会議テンプレートがサーバーに登録されていない可能性があります。</span><span class="sxs-lookup"><span data-stu-id="b9bea-133">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="b9bea-134">詳細については、「 [オンライン会議テンプレートを登録](#register-your-online-meeting-template) する」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="b9bea-134">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="b9bea-135">オンライン会議テンプレートを登録する</span><span class="sxs-lookup"><span data-stu-id="b9bea-135">Register your online-meeting template</span></span>

<span data-ttu-id="b9bea-136">サービスのオンライン会議テンプレートを登録する場合は、詳細情報を含む GitHub の問題を作成できます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-136">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="b9bea-137">その後、登録タイムラインを調整するためにお客様にご連絡します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-137">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="b9bea-138">この記事の最後にある「 **フィードバック** 」セクションに移動します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-138">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="b9bea-139">[ **このページ]** リンクを押します。</span><span class="sxs-lookup"><span data-stu-id="b9bea-139">Press the **This page** link.</span></span>
1. <span data-ttu-id="b9bea-140">新しい問題の **タイトル** を [サービス名] に置き換えて、"my service のオンライン会議テンプレートを登録してください" に設定し `my-service` ます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-140">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="b9bea-141">`newBody`この記事で前述した「[オンライン会議の詳細を実装](#implement-adding-online-meeting-details)する」セクションに記載されている、または同様の変数に設定した文字列で、問題の本文に "[フィードバックをここに入力]" という文字列を置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-141">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="b9bea-142">[ **新しい懸案事項の提出**] をクリックします。</span><span class="sxs-lookup"><span data-stu-id="b9bea-142">Click **Submit new issue**.</span></span>

![Contoso 社のサンプルコンテンツを含む新しい GitHub の問題画面のスクリーンショット](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="b9bea-144">利用可能な Api</span><span class="sxs-lookup"><span data-stu-id="b9bea-144">Available APIs</span></span>

<span data-ttu-id="b9bea-145">この機能では、次の Api を使用できます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-145">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="b9bea-146">予定の開催者 Api</span><span class="sxs-lookup"><span data-stu-id="b9bea-146">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="b9bea-147">[Office. アイテムの件名](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-148">。[開始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start)(時刻) ([時間](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-149">(時刻)-[終了](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end)([時間](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-150">[Office. メールボックス. アイテムの場所](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([場所](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-151">[Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) -任意の出席者 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-152">[Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ....../の内容 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-153">[GetAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-)(body, [Body, setasync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-)) (添付[コンテンツ](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body)の添付)</span><span class="sxs-lookup"><span data-stu-id="b9bea-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="b9bea-154">[CustomProperties (](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)) のようにし[ます。](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-)</span><span class="sxs-lookup"><span data-stu-id="b9bea-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="b9bea-155">[RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="b9bea-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="b9bea-156">認証フローを処理する</span><span class="sxs-lookup"><span data-stu-id="b9bea-156">Handle auth flow</span></span>
  - [<span data-ttu-id="b9bea-157">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="b9bea-157">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="b9bea-158">制限</span><span class="sxs-lookup"><span data-stu-id="b9bea-158">Restrictions</span></span>

<span data-ttu-id="b9bea-159">いくつかの制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-159">Several restrictions apply.</span></span>

- <span data-ttu-id="b9bea-160">オンライン会議サービスプロバイダーにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-160">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="b9bea-161">現時点では、Android はサポートされている唯一のクライアントです。</span><span class="sxs-lookup"><span data-stu-id="b9bea-161">At present, Android is the only supported client.</span></span> <span data-ttu-id="b9bea-162">IOS でのサポートは近日に予定されています。</span><span class="sxs-lookup"><span data-stu-id="b9bea-162">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="b9bea-163">既定の Teams または Skype オプションを置き換えて、管理者によってインストールされたアドインのみが会議の作成画面に表示されます。</span><span class="sxs-lookup"><span data-stu-id="b9bea-163">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="b9bea-164">ユーザーがインストールしたアドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="b9bea-164">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="b9bea-165">アドインアイコンは、16進コード `#919191` または [その他の色の形式](https://convertingcolors.com/hex-color-919191.html)の同等機能を使用したグレースケールである必要があります。</span><span class="sxs-lookup"><span data-stu-id="b9bea-165">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="b9bea-166">予定の開催者 (新規作成) モードでは、UI レスコマンドは1つだけサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b9bea-166">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="b9bea-167">こちらもご覧ください</span><span class="sxs-lookup"><span data-stu-id="b9bea-167">See also</span></span>

- [<span data-ttu-id="b9bea-168">Outlook Mobile のアドイン</span><span class="sxs-lookup"><span data-stu-id="b9bea-168">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="b9bea-169">Outlook Mobile のアドイン コマンドのサポートを追加する</span><span class="sxs-lookup"><span data-stu-id="b9bea-169">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
