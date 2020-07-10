---
title: Outlook Mobile の Outlook のアドイン
description: Outlook mobile アドインは、すべての Microsoft 365 ビジネスアカウントでサポートされており、Outlook.com accounts および support は近日中に gmail アカウントに提供されます。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 34fbb01d596c4da38fe81438088cd71d8c7e152a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093897"
---
# <a name="add-ins-for-outlook-mobile"></a>Outlook Mobile のアドイン

Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.

Outlook mobile アドインは、すべての Microsoft 365 ビジネスアカウントでサポートされており、Outlook.com accounts および support は近日中に Gmail アカウントに提供されます。

**Outlook on iOS の作業ウィンドウの例**

![Outlook on iOS の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane.png)

<br/>

**Outlook on Android の作業ウィンドウの例**

![Outlook on Android の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> アドインは、モバイルブラウザーのモダンバージョンの Outlook では動作しません。 詳細については、「 [Outlook on your mobile browser がアップグレードさ](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)れています。」を参照してください。

## <a name="whats-different-on-mobile"></a>モバイルにおける違い

- The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.
    - アドインは [UI ガイドライン](outlook-addin-design.md)に準拠**していなければなりません**。
    - アドインのシナリオは、[モバイルに対して適切](#what-makes-a-good-scenario-for-mobile-add-ins)である**必要**があります。

- 一般的に、メッセージの読み取りモードのみがサポートされています。 これ `MobileMessageReadCommandSurface` は、マニフェストのモバイルセクションで宣言する必要がある唯一の[extensionpoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface)です。 ただし、予定の開催者モードは、オンライン会議プロバイダー統合アドインでサポートされており、代わりに[MobileOnlineMeetingCommandSurface 拡張点](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)を宣言します。 このシナリオの詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](online-meeting.md)する」の記事を参照してください。

- The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).

- マニフェストで [MobileFormFactor](../reference/manifest/mobileformfactor.md) を使用してストアにアドインを送信する場合、iOS のアドインに関する当社の開発者補遺に同意し、確認のため Apple の開発者 ID を送信しなければなりません。

- 最後に、マニフェストで `MobileFormFactor` を宣言し、適切な種類の[コントロール](../reference/manifest/control.md)と[アイコンのサイズ](../reference/manifest/icon.md)を含める必要があります。

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a>モバイル アドインに対して優れたシナリオにするには

Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.

Outlook Mobile に対して適切なシナリオの例を次に示します。

- The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.

- The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.

**iOS で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![iOS の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

**Android で電子メール メッセージから Trello カードを作成するユーザーの操作の例**

![Android の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a>モバイル上でのアドインのテスト

To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.

After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.

モバイルでのトラブルシューティングは、使用しているツールを持っていない可能性があるため、困難な場合があります。 ただし、iOS でトラブルシューティングを行う方法の1つとして、Fiddler を使用する方法があります ( [ios デバイスでの使用につい](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)ては、このチュートリアルをご覧ください)。

## <a name="next-steps"></a>次の手順

方法はこちら: 

- [モバイル サポートをアドインのマニフェストに追加する](add-mobile-support.md)。
- [アドインで優れたモバイル エクスペリエンスを設計する](outlook-addin-design.md)。
- アドインから[アクセス トークンを取得し、Outlook REST API を呼び出す](use-rest-api.md)。
