---
title: 新規作成フォーム用の Outlook アドインを作成する
description: 新規作成フォーム用の Outlook アドインのシナリオと機能について説明します。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 9156f2e1393c27eea359a6b63da47bc24a8a6828
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234255"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a><span data-ttu-id="e09bc-103">新規作成フォーム用の Outlook アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="e09bc-103">Create Outlook add-ins for compose forms</span></span>

<span data-ttu-id="e09bc-104">Office アドイン マニフェスト用スキーマのバージョン 1.1 と Office.js の v1.1 以降、新規作成アドインを作成することができます。これは、新規作成フォームでアクティブ化される Outlook アドインです。</span><span class="sxs-lookup"><span data-stu-id="e09bc-104">Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of Office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms.</span></span> <span data-ttu-id="e09bc-105">読み取り用アドイン (ユーザーがメッセージまたは予定を閲覧しているときに、読み取りモードでアクティブ化される Outlook アドイン) とは対照的に、新規作成アドインは以下のユーザー シナリオで利用可能です。</span><span class="sxs-lookup"><span data-stu-id="e09bc-105">In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios:</span></span>

- <span data-ttu-id="e09bc-106">新しいメッセージ、会議出席依頼または予定を新規作成フォームで作成している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-106">Composing a new message, meeting request, or appointment in a compose form.</span></span>

- <span data-ttu-id="e09bc-107">既存の予定またはユーザーが開催者になっている会議アイテムを表示または編集している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-107">Viewing or editing an existing appointment, or meeting item in which the user is the organizer.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="e09bc-108">ユーザーが Outlook 2013 および Exchange 2013 の RTM リリースを使用していて、ユーザーが開催する会議アイテムを表示している場合は、使用可能な閲覧アドインを検索できます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-108">If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available.</span></span> <span data-ttu-id="e09bc-109">Office 2013 SP1 リリース以降では、同じシナリオにおいて作成アドインのみをアクティブ化して使用できるような変更が行われています。</span><span class="sxs-lookup"><span data-stu-id="e09bc-109">Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.</span></span>

- <span data-ttu-id="e09bc-110">インライン応答メッセージを作成しているか、別の新規作成フォームでメッセージに返信している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-110">Composing an inline response message or replying to a message in a separate compose form.</span></span>

- <span data-ttu-id="e09bc-111">会議出席依頼または会議アイテムに対する応答 ([**承諾**]、[**仮承諾**]、[**辞退**]) を編集している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-111">Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.</span></span>

- <span data-ttu-id="e09bc-112">会議アイテム用に新しい時間を提案している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-112">Proposing a new time for a meeting item.</span></span>

- <span data-ttu-id="e09bc-113">会議出席依頼や会議アイテムを転送するか、それらに返信している。</span><span class="sxs-lookup"><span data-stu-id="e09bc-113">Forwarding or replying to a meeting request or meeting item.</span></span>

<span data-ttu-id="e09bc-114">これらの各新規作成シナリオでは、アドインで定義されているコマンド ボタンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-114">In each of these compose scenarios, any add-in command buttons defined by the add-in are shown.</span></span> <span data-ttu-id="e09bc-115">アドイン コマンドを実装していない古いアドインでは、ユーザーはリボンにある **Office アドイン** を選択してアドイン選択ウィンドウを開き、新規作成アドインを選択して開始することができます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-115">For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in.</span></span> <span data-ttu-id="e09bc-116">次の図は、新規作成フォームにおけるアドイン コマンドを示しています。</span><span class="sxs-lookup"><span data-stu-id="e09bc-116">The following figure shows add-in commands in a compose form.</span></span>

![アドイン コマンドが含まれた Outlook 作成フォームが表示されています。](../images/compose-form-commands.png)

<span data-ttu-id="e09bc-118">次の図は、ユーザーが Outlook でインライン応答を作成するときにアクティブ化される、アドイン コマンドが実装されていない 2 つの新規作成アドインが含まれたアドイン選択ウィンドウを示しています。</span><span class="sxs-lookup"><span data-stu-id="e09bc-118">The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.</span></span>

![作成されたアイテムに対してアクティブになるテンプレート メール アプリ](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a><span data-ttu-id="e09bc-120">新規作成モードで使用できるアドインの種類</span><span class="sxs-lookup"><span data-stu-id="e09bc-120">Types of add-ins available in compose mode</span></span>

<span data-ttu-id="e09bc-121">新規作成アドインは [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)として実装されます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-121">Compose add-ins are implemented as [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span> <span data-ttu-id="e09bc-122">メール作成または会議出席依頼の返信用のアドインをアクティブ化するために、アドインのマニフェストには [MessageComposeCommandSurface 拡張点要素](../reference/manifest/extensionpoint.md#messagecomposecommandsurface)が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-122">To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) in the manifest.</span></span> <span data-ttu-id="e09bc-123">ユーザーが開催者である予定や会議の新規作成または編集を行うためのアドインをアクティブ化する場合、アドインには [AppointmentOrganizerCommandSurface 拡張点要素](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface)が含まれます。</span><span class="sxs-lookup"><span data-stu-id="e09bc-123">To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span></span>

> [!NOTE]
> <span data-ttu-id="e09bc-124">アドイン コマンドがサポートされていないクライアントまたはサーバー用に開発されたアドインは、[OfficeApp](../reference/manifest/officeapp.md) 要素に含まれる[ルール](../reference/manifest/rule.md)要素の中の[アクティブ化ルール](activation-rules.md)を使用します。</span><span class="sxs-lookup"><span data-stu-id="e09bc-124">Add-ins developed for servers or clients that do not support add-in commands use [activation rules](activation-rules.md) in a [Rule](../reference/manifest/rule.md) element contained in the [OfficeApp](../reference/manifest/officeapp.md) element.</span></span> <span data-ttu-id="e09bc-125">アドインが特に古いクライアントやサーバーのために開発されている場合を除き、新規アドインはアドイン コマンドを使用すべきです。</span><span class="sxs-lookup"><span data-stu-id="e09bc-125">Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.</span></span>

## <a name="api-features-available-to-compose-add-ins"></a><span data-ttu-id="e09bc-126">新規作成アドインに使用できる API の機能</span><span class="sxs-lookup"><span data-stu-id="e09bc-126">API features available to compose add-ins</span></span>

- [<span data-ttu-id="e09bc-127">Outlook で新規作成フォームのアイテムに添付ファイルを追加および削除する</span><span class="sxs-lookup"><span data-stu-id="e09bc-127">Add and remove attachments to an item in a compose form in Outlook</span></span>](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [<span data-ttu-id="e09bc-128">Outlook で新規作成フォームのアイテム データを取得および設定する</span><span class="sxs-lookup"><span data-stu-id="e09bc-128">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="e09bc-129">Outlook の予定またはメッセージを作成するときに受信者を取得、設定、追加する</span><span class="sxs-lookup"><span data-stu-id="e09bc-129">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)
- [<span data-ttu-id="e09bc-130">Outlook で予定またはメッセージを作成するときに件名を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="e09bc-130">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="e09bc-131">Outlook で予定またはメッセージを作成するときに本文にデータを挿入する</span><span class="sxs-lookup"><span data-stu-id="e09bc-131">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="e09bc-132">Outlook で予定を作成するときに場所を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="e09bc-132">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="e09bc-133">Outlook で予定を作成するときに時刻を取得または設定する</span><span class="sxs-lookup"><span data-stu-id="e09bc-133">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a><span data-ttu-id="e09bc-134">関連項目</span><span class="sxs-lookup"><span data-stu-id="e09bc-134">See also</span></span>

- [<span data-ttu-id="e09bc-135">Office の Outlook アドインの概要</span><span class="sxs-lookup"><span data-stu-id="e09bc-135">Get Started with Outlook add-ins for Office</span></span>](../quickstarts/outlook-quickstart.md)
