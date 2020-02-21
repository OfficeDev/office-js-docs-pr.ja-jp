---
title: Outlook Mobile の Outlook のアドイン
description: Outlook Mobile アドインはすべての商用版 Office 365 アカウント、Outlook.com アカウントでサポートされ、近いうちに Gmail アカウントでもサポートされる予定です。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 7ede3165f40e644715dc488214e047f00dafbede
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166555"
---
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="0b43e-103">Outlook Mobile のアドイン</span><span class="sxs-lookup"><span data-stu-id="0b43e-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="0b43e-p101">現時点で、アドインは他の Outlook エンドポイントで利用できるものと同じ API を使用して Outlook Mobile で動作します。Outlook 用のアドインを作成済みの場合、簡単に Outlook Mobile で動作するようにできます。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p101">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints. If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="0b43e-106">Outlook Mobile アドインはすべての商用版 Office 365 アカウント、Outlook.com アカウントでサポートされ、近いうちに Gmail アカウントでもサポートされる予定です。</span><span class="sxs-lookup"><span data-stu-id="0b43e-106">Outlook mobile add-ins are supported on all Office 365 Commercial accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="0b43e-107">**Outlook on iOS の作業ウィンドウの例**</span><span class="sxs-lookup"><span data-stu-id="0b43e-107">**An example task pane in Outlook on iOS**</span></span>

![Outlook on iOS の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="0b43e-109">**Outlook on Android の作業ウィンドウの例**</span><span class="sxs-lookup"><span data-stu-id="0b43e-109">**An example task pane in Outlook on Android**</span></span>

![Outlook on Android の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane-android.png)

## <a name="whats-different-on-mobile"></a><span data-ttu-id="0b43e-111">モバイルにおける違い</span><span class="sxs-lookup"><span data-stu-id="0b43e-111">What's different on mobile?</span></span>

- <span data-ttu-id="0b43e-p102">モバイル用の設計において、小さいサイズと迅速な操作性が課題となります。お客様に高品質のエクスペリエンスを提供するため、モバイル サポートを宣言するアドインに対して厳格な検証条件を定めています。AppSource で承認を得るには、この条件を満たす必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p102">The small size and quick interactions make designing for mobile a challenge. To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="0b43e-114">アドインは [UI ガイドライン](outlook-addin-design.md)に準拠**していなければなりません**。</span><span class="sxs-lookup"><span data-stu-id="0b43e-114">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="0b43e-115">アドインのシナリオは、[モバイルに対して適切](#what-makes-a-good-scenario-for-mobile-add-ins)である**必要**があります。</span><span class="sxs-lookup"><span data-stu-id="0b43e-115">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="0b43e-p103">現時点では、メールの読み取りのみがサポートされています。つまり、`MobileMessageReadCommandSurface` は、マニフェストのモバイル セクションで宣言する必要がある唯一の [ExtensionPoint](../reference/manifest/extensionpoint.md) になります。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p103">Only mail read is supported at this time. That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md) you should declare in the mobile section of your manifest.</span></span>

- <span data-ttu-id="0b43e-p104">[makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API はモバイルではサポートされていません。モバイル アプリは REST API を使用して、サーバーと通信します。アプリのバックエンドで Exchange サーバーと接続する必要がある場合、コールバック トークンを使用して REST API 呼び出しを行うことができます。詳しくは、「[Outlook アドインからの Outlook REST API の使用](use-rest-api.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p104">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server. If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls. For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="0b43e-121">マニフェストで [MobileFormFactor](../reference/manifest/mobileformfactor.md) を使用してストアにアドインを送信する場合、iOS のアドインに関する当社の開発者補遺に同意し、確認のため Apple の開発者 ID を送信しなければなりません。</span><span class="sxs-lookup"><span data-stu-id="0b43e-121">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="0b43e-122">最後に、マニフェストで `MobileFormFactor` を宣言し、適切な種類の[コントロール](../reference/manifest/control.md)と[アイコンのサイズ](../reference/manifest/icon.md)を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b43e-122">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="0b43e-123">モバイル アドインに対して優れたシナリオにするには</span><span class="sxs-lookup"><span data-stu-id="0b43e-123">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="0b43e-p105">電話での Outlook セッションの平均の長さは PC よりも短いことを忘れないでください。つまり、アドインを高速にする必要があります。さらに、シナリオでは、ユーザーの電子メール フローに出入りし、中断せずに続行できるようにする必要もあります。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p105">Remember that the average Outlook session length on a phone is much shorter than on a PC. That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="0b43e-126">Outlook Mobile に対して適切なシナリオの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="0b43e-126">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="0b43e-p106">アドインを使用すると、貴重な情報を Outlook に伝えることができるため、ユーザーは電子メールをトリアージし、適切に対応できます。例: ユーザーが顧客情報を確認し、適切な情報を共有するための CRM アドイン。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p106">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately. Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="0b43e-p107">アドインが、追跡システム、共同作業システム、または類似するシステムに情報を保存して、ユーザーの電子メール コンテンツに価値を追加します。例: ユーザーが電子メールを、プロジェクト進捗管理用にタスク項目に変換したり、サポート チーム用にヘルプ チケットに変換したりするアドイン。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p107">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system. Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="0b43e-131">**iOS で電子メール メッセージから Trello カードを作成するユーザーの操作の例**</span><span class="sxs-lookup"><span data-stu-id="0b43e-131">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![iOS の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="0b43e-133">**Android で電子メール メッセージから Trello カードを作成するユーザーの操作の例**</span><span class="sxs-lookup"><span data-stu-id="0b43e-133">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Android の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="0b43e-135">モバイル上でのアドインのテスト</span><span class="sxs-lookup"><span data-stu-id="0b43e-135">Testing your add-ins on mobile</span></span>

<span data-ttu-id="0b43e-p108">Outlook Mobile でアドインをテストするために、O365 や Outlook.com アカウントにアドインをサイドローディングできます。Outlook on the web で、設定ギアに移動し、[**統合の管理**] または [**アドインの管理**] を選択します。上部付近で、[**カスタム アドインを追加するには、ここをクリックします**] をクリックし、マニフェストをアップロードします。マニフェストの形式に `MobileFormFactor` が含まれていることを確認します。含まれていないと、読み込むことができません。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p108">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account. In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest. Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="0b43e-p109">アドインが動作することを確認したら、携帯電話やタブレットなど、別のサイズの画面でテストします。コンストラストやフォント サイズ、色、さらには VoiceOver (iOS) または TalkBack (Android) などのスクリーン リーダーが使用できることなど、アクセシビリティのガイドラインに従っていることも確認してください。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p109">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets. You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="0b43e-p110">モバイルにおけるトラブルシューティングは、使い慣れたツールがないことがあるため難しい場合があります。トラブルシューティングの 1 つのオプションは、[Vorlon.js を使用](../testing/debug-office-add-ins-on-ipad-and-mac.md)する方法です。または、Fiddler を以前に使用したことがある場合、[iOS デバイスでの使用についてはこのチュートリアル](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)をご確認ください。</span><span class="sxs-lookup"><span data-stu-id="0b43e-p110">Troubleshooting on mobile can be hard since you may not have the tools you're used to. One option for troubleshooting is to [use Vorlon.js](../testing/debug-office-add-ins-on-ipad-and-mac.md). Or, if you've used Fiddler before, check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices).</span></span>

## <a name="next-steps"></a><span data-ttu-id="0b43e-144">次の手順</span><span class="sxs-lookup"><span data-stu-id="0b43e-144">Next steps</span></span>

<span data-ttu-id="0b43e-145">方法はこちら: </span><span class="sxs-lookup"><span data-stu-id="0b43e-145">Learn how to:</span></span>

- <span data-ttu-id="0b43e-146">[モバイル サポートをアドインのマニフェストに追加する](add-mobile-support.md)。</span><span class="sxs-lookup"><span data-stu-id="0b43e-146">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="0b43e-147">[アドインで優れたモバイル エクスペリエンスを設計する](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="0b43e-147">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="0b43e-148">アドインから[アクセス トークンを取得し、Outlook REST API を呼び出す](use-rest-api.md)。</span><span class="sxs-lookup"><span data-stu-id="0b43e-148">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
