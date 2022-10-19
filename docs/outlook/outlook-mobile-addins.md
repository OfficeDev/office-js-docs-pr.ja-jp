---
title: Outlook Mobile の Outlook のアドイン
description: Outlook モバイル アドインは、すべての Microsoft 365 ビジネス アカウントと Outlook.com アカウントでサポートされています。
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: ca09ba550d8d2ed6e9003e85a8d042f413a6ab52
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607563"
---
# <a name="add-ins-for-outlook-mobile"></a>Outlook Mobile のアドイン

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Outlook モバイル アドインは、すべての Microsoft 365 ビジネス アカウントと Outlook.com アカウントでサポートされています。 ただし、現在、Gmail アカウントではサポートを利用できません。

**Outlook on iOS の作業ウィンドウの例**

![Outlook on iOS の作業ウィンドウのスクリーンショット。](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Outlook on Android の作業ウィンドウの例**

![Android 上の Outlook の作業ウィンドウのスクリーンショット。](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a>モバイルにおける違い

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
  - アドインは [UI ガイドライン](outlook-addin-design.md)に準拠 **していなければなりません**。
  - アドインのシナリオは、[モバイルに対して適切](#what-makes-a-good-scenario-for-mobile-add-ins)である **必要** があります。

[!INCLUDE [Teams manifest not supported on mobile devices](../includes/no-mobile-with-json-note.md)]

- 一般に、現時点では、メッセージ読み取りモードのみがサポートされています。 つまり、マニフェストの `MobileMessageReadCommandSurface` モバイル セクションで宣言する必要がある唯一の [ExtensionPoint](/javascript/api/manifest/extensionpoint#mobilemessagereadcommandsurface) です。 ただし、いくつかの例外があります。
  1. 予定開催者モードは、 [MobileOnlineMeetingCommandSurface 拡張ポイント](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface)を宣言するオンライン会議プロバイダー統合アドインでサポートされています。 このシナリオの詳細については、 [オンライン会議プロバイダーの Outlook モバイル アドインの作成](online-meeting.md) に関する記事を参照してください。
  1. 予定出席者モードは、メモと顧客関係管理 (CRM) アプリケーションのプロバイダーによって作成された統合アドインでサポートされます。 このようなアドインは、代わりに [MobileLogEventAppointmentAttendee 拡張ポイント](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee)を宣言する必要があります。 このシナリオの詳細については、 [Outlook モバイル アドインの外部アプリケーションに対する予定ノートのログ](mobile-log-appointments.md) 記録に関する記事を参照してください。

- The [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- マニフェストで [MobileFormFactor](/javascript/api/manifest/mobileformfactor) を使用してストアにアドインを送信する場合、iOS のアドインに関する当社の開発者補遺に同意し、確認のため Apple の開発者 ID を送信しなければなりません。

- 最後に、マニフェストで `MobileFormFactor` を宣言し、適切な種類の[コントロール](/javascript/api/manifest/control)と[アイコンのサイズ](/javascript/api/manifest/icon)を含める必要があります。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>モバイル アドインに対して優れたシナリオにするには

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Outlook Mobile に対して適切なシナリオの例を次に示します。

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**iOS で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![iOS 上の Outlook Mobile アドインとのユーザー操作を示すアニメーション GIF。](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Android で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![Android 上の Outlook Mobile アドインとのユーザー操作を示すアニメーション GIF。](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>モバイル上でのアドインのテスト

Outlook Mobile でアドインをテストするには、まずアドインを Microsoft 365 または Web、Windows、または Mac 上の Outlook.com アカウントに [サイドロード](sideload-outlook-add-ins-for-testing.md) します。 マニフェストが適切に書式設定されて含まれている `MobileFormFactor` か、モバイル上の Outlook クライアントに読み込まれないことを確認します。

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

使い慣れているツールがない可能性があるため、モバイルでのトラブルシューティングは難しい場合があります。 ただし、iOS でのトラブルシューティングの 1 つのオプションは、Fiddler を使用することです ( [iOS デバイスでの使用に関するこのチュートリアルを](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)参照してください)。

> [!NOTE]
> iPhone および Android スマートフォンの最新のOutlook on the webは、Outlook アドインのテストに必要または利用できなくなりました。さらに、アドインは、Outlook on Android、iOS、およびオンプレミスの Exchange アカウントを使用した最新のモバイル Web ではサポートされていません。 一部の iOS デバイスでは、従来のOutlook on the webでオンプレミスの Exchange アカウントを使用する場合でもアドインがサポートされます。 サポートされているブラウザーの詳細については、「[Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md#client-requirements-non-windows-smartphone-and-tablet)」を参照してください。

## <a name="next-steps"></a>次の手順

方法はこちら: 

- [モバイル サポートをアドインのマニフェストに追加する](add-mobile-support.md)。
- [アドインで優れたモバイル エクスペリエンスを設計する](outlook-addin-design.md)。
- アドインから[アクセス トークンを取得し、Outlook REST API を呼び出す](use-rest-api.md)。
