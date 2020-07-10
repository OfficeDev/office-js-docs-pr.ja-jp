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
# <a name="add-ins-for-outlook-mobile"></a><span data-ttu-id="9aac8-103">Outlook Mobile のアドイン</span><span class="sxs-lookup"><span data-stu-id="9aac8-103">Add-ins for Outlook Mobile</span></span>

<span data-ttu-id="9aac8-104">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints.</span><span class="sxs-lookup"><span data-stu-id="9aac8-104">Add-ins now work on Outlook Mobile, using the same APIs available for other Outlook endpoints.</span></span> <span data-ttu-id="9aac8-105">If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="9aac8-105">If you've built an add-in for Outlook already, it's easy to get it working on Outlook Mobile.</span></span>

<span data-ttu-id="9aac8-106">Outlook mobile アドインは、すべての Microsoft 365 ビジネスアカウントでサポートされており、Outlook.com accounts および support は近日中に Gmail アカウントに提供されます。</span><span class="sxs-lookup"><span data-stu-id="9aac8-106">Outlook mobile add-ins are supported on all Microsoft 365 business accounts, Outlook.com accounts, and support is coming soon to Gmail accounts.</span></span>

<span data-ttu-id="9aac8-107">**Outlook on iOS の作業ウィンドウの例**</span><span class="sxs-lookup"><span data-stu-id="9aac8-107">**An example task pane in Outlook on iOS**</span></span>

![Outlook on iOS の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane.png)

<br/>

<span data-ttu-id="9aac8-109">**Outlook on Android の作業ウィンドウの例**</span><span class="sxs-lookup"><span data-stu-id="9aac8-109">**An example task pane in Outlook on Android**</span></span>

![Outlook on Android の作業ウィンドウのスクリーンショット](../images/outlook-mobile-addin-taskpane-android.png)

> [!IMPORTANT]
> <span data-ttu-id="9aac8-111">アドインは、モバイルブラウザーのモダンバージョンの Outlook では動作しません。</span><span class="sxs-lookup"><span data-stu-id="9aac8-111">Add-ins don't work in the modern version of Outlook in a mobile browser.</span></span> <span data-ttu-id="9aac8-112">詳細については、「 [Outlook on your mobile browser がアップグレードさ](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816)れています。」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9aac8-112">For more information, see [Outlook on your mobile browser is being upgraded](https://techcommunity.microsoft.com/t5/outlook-blog/outlook-on-your-mobile-browser-is-being-upgraded/ba-p/1125816).</span></span>

## <a name="whats-different-on-mobile"></a><span data-ttu-id="9aac8-113">モバイルにおける違い</span><span class="sxs-lookup"><span data-stu-id="9aac8-113">What's different on mobile?</span></span>

- <span data-ttu-id="9aac8-114">The small size and quick interactions make designing for mobile a challenge.</span><span class="sxs-lookup"><span data-stu-id="9aac8-114">The small size and quick interactions make designing for mobile a challenge.</span></span> <span data-ttu-id="9aac8-115">To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span><span class="sxs-lookup"><span data-stu-id="9aac8-115">To ensure quality experiences for our customers, we are setting strict validation criteria that must be met by an add-in declaring mobile support, in order to be approved in AppSource.</span></span>
    - <span data-ttu-id="9aac8-116">アドインは [UI ガイドライン](outlook-addin-design.md)に準拠**していなければなりません**。</span><span class="sxs-lookup"><span data-stu-id="9aac8-116">The add-in **MUST** adhere to the [UI guidelines](outlook-addin-design.md).</span></span>
    - <span data-ttu-id="9aac8-117">アドインのシナリオは、[モバイルに対して適切](#what-makes-a-good-scenario-for-mobile-add-ins)である**必要**があります。</span><span class="sxs-lookup"><span data-stu-id="9aac8-117">The scenario for the add-in **MUST** [make sense on mobile](#what-makes-a-good-scenario-for-mobile-add-ins).</span></span>

- <span data-ttu-id="9aac8-118">一般的に、メッセージの読み取りモードのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="9aac8-118">In general, only Message Read mode is supported at this time.</span></span> <span data-ttu-id="9aac8-119">これ `MobileMessageReadCommandSurface` は、マニフェストのモバイルセクションで宣言する必要がある唯一の[extensionpoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface)です。</span><span class="sxs-lookup"><span data-stu-id="9aac8-119">That means `MobileMessageReadCommandSurface` is the only [ExtensionPoint](../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface) you should declare in the mobile section of your manifest.</span></span> <span data-ttu-id="9aac8-120">ただし、予定の開催者モードは、オンライン会議プロバイダー統合アドインでサポートされており、代わりに[MobileOnlineMeetingCommandSurface 拡張点](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)を宣言します。</span><span class="sxs-lookup"><span data-stu-id="9aac8-120">However, Appointment Organizer mode is supported for online meeting provider integrated add-ins which instead declare the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview).</span></span> <span data-ttu-id="9aac8-121">このシナリオの詳細については、「[オンライン会議プロバイダー用の Outlook モバイルアドインを作成](online-meeting.md)する」の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="9aac8-121">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this scenario.</span></span>

- <span data-ttu-id="9aac8-122">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server.</span><span class="sxs-lookup"><span data-stu-id="9aac8-122">The [makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) API is not supported on mobile since the mobile app uses REST APIs to communicate with the server.</span></span> <span data-ttu-id="9aac8-123">If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls.</span><span class="sxs-lookup"><span data-stu-id="9aac8-123">If your app backend needs to connect to the Exchange server, you can use the callback token to make REST API calls.</span></span> <span data-ttu-id="9aac8-124">For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span><span class="sxs-lookup"><span data-stu-id="9aac8-124">For details, see [Use the Outlook REST APIs from an Outlook add-in](use-rest-api.md).</span></span>

- <span data-ttu-id="9aac8-125">マニフェストで [MobileFormFactor](../reference/manifest/mobileformfactor.md) を使用してストアにアドインを送信する場合、iOS のアドインに関する当社の開発者補遺に同意し、確認のため Apple の開発者 ID を送信しなければなりません。</span><span class="sxs-lookup"><span data-stu-id="9aac8-125">When you submit your add-in to the store with [MobileFormFactor](../reference/manifest/mobileformfactor.md) in the manifest, you'll need to agree to our developer addendum for add-ins on iOS, and you must submit your Apple Developer ID for verification.</span></span>

- <span data-ttu-id="9aac8-126">最後に、マニフェストで `MobileFormFactor` を宣言し、適切な種類の[コントロール](../reference/manifest/control.md)と[アイコンのサイズ](../reference/manifest/icon.md)を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="9aac8-126">Finally, your manifest will need to declare `MobileFormFactor`, and have the correct types of [controls](../reference/manifest/control.md) and [icon sizes](../reference/manifest/icon.md) included.</span></span>

## <a name="what-makes-a-good-scenario-for-mobile-add-ins"></a><span data-ttu-id="9aac8-127">モバイル アドインに対して優れたシナリオにするには</span><span class="sxs-lookup"><span data-stu-id="9aac8-127">What makes a good scenario for mobile add-ins?</span></span>

<span data-ttu-id="9aac8-128">Remember that the average Outlook session length on a phone is much shorter than on a PC.</span><span class="sxs-lookup"><span data-stu-id="9aac8-128">Remember that the average Outlook session length on a phone is much shorter than on a PC.</span></span> <span data-ttu-id="9aac8-129">That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span><span class="sxs-lookup"><span data-stu-id="9aac8-129">That means your add-in must be fast, and the scenario must allow the user to get in, get out, and get on with their email workflow.</span></span>

<span data-ttu-id="9aac8-130">Outlook Mobile に対して適切なシナリオの例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="9aac8-130">Here are examples of scenarios that make sense in Outlook Mobile.</span></span>

- <span data-ttu-id="9aac8-131">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately.</span><span class="sxs-lookup"><span data-stu-id="9aac8-131">The add-in brings valuable information into Outlook, helping users triage their email and respond appropriately.</span></span> <span data-ttu-id="9aac8-132">Example: a CRM add-in that lets the user see customer information and share appropriate information.</span><span class="sxs-lookup"><span data-stu-id="9aac8-132">Example: a CRM add-in that lets the user see customer information and share appropriate information.</span></span>

- <span data-ttu-id="9aac8-133">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system.</span><span class="sxs-lookup"><span data-stu-id="9aac8-133">The add-in adds value to the user's email content by saving the information to a tracking, collaboration, or similar system.</span></span> <span data-ttu-id="9aac8-134">Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span><span class="sxs-lookup"><span data-stu-id="9aac8-134">Example: an add-in that lets users turn emails into task items for project tracking, or help tickets for a support team.</span></span>

<span data-ttu-id="9aac8-135">**iOS で電子メール メッセージから Trello カードを作成するユーザーの操作の例**</span><span class="sxs-lookup"><span data-stu-id="9aac8-135">**An example user interaction to create a Trello card from an email message on iOS**</span></span>

![iOS の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction.gif)

<br/>

<span data-ttu-id="9aac8-137">**Android で電子メール メッセージから Trello カードを作成するユーザーの操作の例**</span><span class="sxs-lookup"><span data-stu-id="9aac8-137">**An example user interaction to create a Trello card from an email message on Android**</span></span>

![Android の Outlook Mobile アドインを使用したユーザーの操作を示すアニメーション GIF](../images/outlook-mobile-addin-interaction-android.gif)

## <a name="testing-your-add-ins-on-mobile"></a><span data-ttu-id="9aac8-139">モバイル上でのアドインのテスト</span><span class="sxs-lookup"><span data-stu-id="9aac8-139">Testing your add-ins on mobile</span></span>

<span data-ttu-id="9aac8-140">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account.</span><span class="sxs-lookup"><span data-stu-id="9aac8-140">To test an add-in on Outlook Mobile, you can sideload an add-in to an O365 or Outlook.com account.</span></span> <span data-ttu-id="9aac8-141">In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest.</span><span class="sxs-lookup"><span data-stu-id="9aac8-141">In Outlook on the web, go to the settings gear, and choose **Manage Integrations** or **Manage Add-ins**. Near the top, click where it says **Click here to add a custom add-in** and upload your manifest.</span></span> <span data-ttu-id="9aac8-142">Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span><span class="sxs-lookup"><span data-stu-id="9aac8-142">Make sure your manifest is properly formatted to contain `MobileFormFactor` or it won't load.</span></span>

<span data-ttu-id="9aac8-143">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets.</span><span class="sxs-lookup"><span data-stu-id="9aac8-143">After your add-in is working, make sure to test it on different screen sizes, including phones and tablets.</span></span> <span data-ttu-id="9aac8-144">You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span><span class="sxs-lookup"><span data-stu-id="9aac8-144">You should make sure it meets accessibility guidelines for contrast, font size, and color, as well as being usable with a screen reader such as VoiceOver on iOS or TalkBack on Android.</span></span>

<span data-ttu-id="9aac8-145">モバイルでのトラブルシューティングは、使用しているツールを持っていない可能性があるため、困難な場合があります。</span><span class="sxs-lookup"><span data-stu-id="9aac8-145">Troubleshooting on mobile can be hard since you may not have the tools you're used to.</span></span> <span data-ttu-id="9aac8-146">ただし、iOS でトラブルシューティングを行う方法の1つとして、Fiddler を使用する方法があります ( [ios デバイスでの使用につい](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)ては、このチュートリアルをご覧ください)。</span><span class="sxs-lookup"><span data-stu-id="9aac8-146">However, one option for troubleshooting on iOS is to use Fiddler (check out [this tutorial on using it with an iOS device](https://www.telerik.com/blogs/using-fiddler-with-apple-ios-devices)).</span></span>

## <a name="next-steps"></a><span data-ttu-id="9aac8-147">次の手順</span><span class="sxs-lookup"><span data-stu-id="9aac8-147">Next steps</span></span>

<span data-ttu-id="9aac8-148">方法はこちら: </span><span class="sxs-lookup"><span data-stu-id="9aac8-148">Learn how to:</span></span>

- <span data-ttu-id="9aac8-149">[モバイル サポートをアドインのマニフェストに追加する](add-mobile-support.md)。</span><span class="sxs-lookup"><span data-stu-id="9aac8-149">[Add mobile support to your add-in's manifest](add-mobile-support.md).</span></span>
- <span data-ttu-id="9aac8-150">[アドインで優れたモバイル エクスペリエンスを設計する](outlook-addin-design.md)。</span><span class="sxs-lookup"><span data-stu-id="9aac8-150">[Design a great mobile experience for your add-in](outlook-addin-design.md).</span></span>
- <span data-ttu-id="9aac8-151">アドインから[アクセス トークンを取得し、Outlook REST API を呼び出す](use-rest-api.md)。</span><span class="sxs-lookup"><span data-stu-id="9aac8-151">[Get an access token and call Outlook REST APIs](use-rest-api.md) from your add-in.</span></span>
