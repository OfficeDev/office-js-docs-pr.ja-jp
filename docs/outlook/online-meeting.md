---
title: オンライン会議プロバイダー用の Outlook モバイルアドインを作成する (プレビュー)
description: オンライン会議サービスプロバイダー用の Outlook mobile アドインをセットアップする方法について説明します。
ms.topic: article
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: 841200d8db1dc4c7a89c953737f0bc5b74edf7ea
ms.sourcegitcommit: 76552b3e5725d9112c772595971b922c295e6b4c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/10/2020
ms.locfileid: "43226030"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider-preview"></a><span data-ttu-id="cbe09-103">オンライン会議プロバイダー用の Outlook モバイルアドインを作成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="cbe09-103">Create an Outlook mobile add-in for an online-meeting provider (preview)</span></span>

<span data-ttu-id="cbe09-104">オンライン会議の設定は、Outlook ユーザーにとって中心的な操作であり、Outlook mobile[を使用して Teams 会議を](/microsoftteams/teams-add-in-for-outlook)簡単に作成できます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="cbe09-105">ただし、Microsoft 以外のサービスを使用して Outlook でオンライン会議を作成するのは煩雑な場合があります。</span><span class="sxs-lookup"><span data-stu-id="cbe09-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="cbe09-106">この機能を実装することにより、サービスプロバイダーは、Outlook アドインユーザーに対してオンライン会議の作成環境を合理化することができます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!NOTE]
> <span data-ttu-id="cbe09-107">この機能は、Office 365 サブスクリプションを使用した Android の[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-107">This feature is only supported in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) on Android with an Office 365 subscription.</span></span>

<span data-ttu-id="cbe09-108">この記事では、ユーザーがオンライン会議サービスを使用して会議を整理し、会議に参加できるようにするために Outlook モバイルアドインをセットアップする方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="cbe09-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="cbe09-109">この記事全体で、架空のオンライン会議サービスプロバイダーである "Contoso" を使用します。</span><span class="sxs-lookup"><span data-stu-id="cbe09-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="cbe09-110">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="cbe09-110">Configure the manifest</span></span>

<span data-ttu-id="cbe09-111">ユーザーがアドインを使用してオンライン会議を作成できるようにするには`MobileOnlineMeetingCommandSurface` 、マニフェストで親要素`MobileFormFactor`の下に拡張点を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbe09-111">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="cbe09-112">その他のフォームファクターはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="cbe09-112">Other form factors are not supported.</span></span>

<span data-ttu-id="cbe09-113">次の例は、 `MobileFormFactor`要素と`MobileOnlineMeetingCommandSurface`拡張点を含むマニフェストのサンプルを示しています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-113">The following example shows a sample of the manifest that includes the `MobileFormFactor` element and `MobileOnlineMeetingCommandSurface` extension point.</span></span>

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

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="cbe09-114">オンライン会議の詳細の追加を実装する</span><span class="sxs-lookup"><span data-stu-id="cbe09-114">Implement adding online meeting details</span></span>

<span data-ttu-id="cbe09-115">このセクションでは、アドインスクリプトでユーザーの会議を更新してオンライン会議の詳細を含める方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="cbe09-115">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

<span data-ttu-id="cbe09-116">次の例は、オンライン会議の詳細を作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-116">The following example shows how you construct online meeting details.</span></span> <span data-ttu-id="cbe09-117">ここでは、サービスから会議開催者の ID とその他の詳細を取得する方法については説明しません。</span><span class="sxs-lookup"><span data-stu-id="cbe09-117">Not shown is how to get the meeting organizer's ID and other details from your service.</span></span>

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

<span data-ttu-id="cbe09-118">次の例は、マニフェストで`insertContosoMeeting`参照される UI を使用しない関数を定義して、オンライン会議の詳細で会議の本文を更新する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-118">The following example shows how to define a UI-less function named `insertContosoMeeting` referenced in the manifest to update the meeting body with the online meeting details.</span></span>

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

<span data-ttu-id="cbe09-119">次の例は、前の例で使用`updateBody`した、会議の現在の本文にオンライン会議の詳細を追加する、サポート関数の実装を示しています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-119">The following example shows an implementation of the supporting function `updateBody` used in the previous example that appends the online meeting details to the current body of the meeting.</span></span>

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

## <a name="testing-and-validation"></a><span data-ttu-id="cbe09-120">テストと検証</span><span class="sxs-lookup"><span data-stu-id="cbe09-120">Testing and validation</span></span>

<span data-ttu-id="cbe09-121">通常のガイダンスに従って、[アドインをテストし、検証](testing-and-tips.md)します。</span><span class="sxs-lookup"><span data-stu-id="cbe09-121">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="cbe09-122">Outlook on the web、Windows、または Mac で[サイド](sideload-outlook-add-ins-for-testing.md)ロード後に、android モバイルデバイスで outlook を再起動します (現時点でサポートされている唯一のクライアントは android です)。</span><span class="sxs-lookup"><span data-stu-id="cbe09-122">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device (Android is the only supported client for now).</span></span> <span data-ttu-id="cbe09-123">次に、新しい会議画面で、Microsoft Teams または Skype のトグルが自分のものに置き換えられていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="cbe09-123">Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="cbe09-124">会議 UI を作成する</span><span class="sxs-lookup"><span data-stu-id="cbe09-124">Create meeting UI</span></span>

<span data-ttu-id="cbe09-125">会議の開催者として、会議を作成するときに次の3つの画像のような画面が表示されます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-125">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="cbe09-126">[[android での会議の作成] 画面のスクリーンショット-contoso の [会議を作成する] 画面のスクリーンショットの作成-contoso-contoso での会議の作成画面のスクリーンショットを切り替える![](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="cbe09-126">[![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="cbe09-127">ミーティング UI への参加</span><span class="sxs-lookup"><span data-stu-id="cbe09-127">Join meeting UI</span></span>

<span data-ttu-id="cbe09-128">会議の出席者として、会議を表示するときに次のような画面が表示されます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-128">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="cbe09-129">[![Android に参加した会議画面のスクリーンショット](../images/outlook-android-join-online-meeting-view.png)](../images/outlook-android-join-online-meeting-view-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="cbe09-129">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view.png)](../images/outlook-android-join-online-meeting-view-expanded.png#lightbox)</span></span>

## <a name="available-apis"></a><span data-ttu-id="cbe09-130">利用可能な Api</span><span class="sxs-lookup"><span data-stu-id="cbe09-130">Available APIs</span></span>

<span data-ttu-id="cbe09-131">この機能では、次の Api を使用できます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-131">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="cbe09-132">予定の開催者 Api</span><span class="sxs-lookup"><span data-stu-id="cbe09-132">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="cbe09-133">[Office. アイテムの件名](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject)([subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-133">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-134">。[開始](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start)(時刻) ([時間](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-134">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-135">(時刻)-[終了](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end)([時間](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-135">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-136">[Office. メールボックス. アイテムの場所](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location)([場所](/javascript/api/outlook/office.location?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-136">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-137">[Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) -任意の出席者 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-137">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-138">「 [Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) ..」 (開催[者](/javascript/api/outlook/office.organizer?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-138">[Office.context.mailbox.item.organizer](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) ([Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-139">[Office.](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ....../の内容 ([受信者](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-139">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-140">[GetAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-)(body, [Body, setasync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-)) (添付[コンテンツ](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body)の添付)</span><span class="sxs-lookup"><span data-stu-id="cbe09-140">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="cbe09-141">[CustomProperties (](/javascript/api/outlook/office.customproperties?view=outlook-js-preview)) のようにし[ます。](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-)</span><span class="sxs-lookup"><span data-stu-id="cbe09-141">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="cbe09-142">[RoamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="cbe09-142">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span></span>
- <span data-ttu-id="cbe09-143">認証フローを処理する</span><span class="sxs-lookup"><span data-stu-id="cbe09-143">Handle auth flow</span></span>
  - [<span data-ttu-id="cbe09-144">ダイアログ API</span><span class="sxs-lookup"><span data-stu-id="cbe09-144">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="cbe09-145">制限</span><span class="sxs-lookup"><span data-stu-id="cbe09-145">Restrictions</span></span>

<span data-ttu-id="cbe09-146">いくつかの制限が適用されます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-146">Several restrictions apply.</span></span>

- <span data-ttu-id="cbe09-147">オンライン会議サービスプロバイダーにのみ適用されます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-147">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="cbe09-148">現在プレビュー中であるため、この機能は運用アドインでは使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="cbe09-148">Currently in preview so this feature shouldn't be used in production add-ins.</span></span>
- <span data-ttu-id="cbe09-149">現時点では、Android はサポートされている唯一のクライアントです。</span><span class="sxs-lookup"><span data-stu-id="cbe09-149">At present, Android is the only supported client.</span></span> <span data-ttu-id="cbe09-150">IOS でのサポートは近日に予定されています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-150">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="cbe09-151">既定の Teams または Skype オプションを置き換えて、管理者によってインストールされたアドインのみが会議の作成画面に表示されます。</span><span class="sxs-lookup"><span data-stu-id="cbe09-151">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="cbe09-152">ユーザーがインストールしたアドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="cbe09-152">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="cbe09-153">アドインアイコンは、16進コード`#919191`または[その他の色の形式](https://convertingcolors.com/hex-color-919191.html)の同等機能を使用したグレースケールである必要があります。</span><span class="sxs-lookup"><span data-stu-id="cbe09-153">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="cbe09-154">予定の開催者 (新規作成) モードでは、UI レスコマンドは1つだけサポートされています。</span><span class="sxs-lookup"><span data-stu-id="cbe09-154">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="cbe09-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="cbe09-155">See also</span></span>

- [<span data-ttu-id="cbe09-156">Outlook Mobile のアドイン</span><span class="sxs-lookup"><span data-stu-id="cbe09-156">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="cbe09-157">Outlook Mobile のアドイン コマンドのサポートを追加する</span><span class="sxs-lookup"><span data-stu-id="cbe09-157">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
