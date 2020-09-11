---
title: Outlook アドインの送信時機能
description: アイテムを処理する方法、またはユーザーが特定のアクションを実行しないようにする方法を提供し、送信時にアドインが特定のプロパティを設定できるようにします。
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: b0c33cda6a03ea7a67ca2eae2bd31416b0794df0
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431053"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="7fd81-103">Outlook アドインの送信時機能</span><span class="sxs-lookup"><span data-stu-id="7fd81-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="7fd81-p101">Outlook アドインの送信時機能は、メッセージまたは会議アイテムを処理する方法、またはユーザーが特定のアクションを実行できないようにする方法を提供し、送信時にアドインが特定のプロパティを設定できるようにします。たとえば、送信時機能を使用すると次のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="7fd81-106">ユーザーが機密情報を送信したり、件名を空白にしたままにしないようにする。</span><span class="sxs-lookup"><span data-stu-id="7fd81-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="7fd81-107">特定の受信者をメッセージの CC 行または会議の任意の受信者行に追加する。</span><span class="sxs-lookup"><span data-stu-id="7fd81-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="7fd81-108">送信時の機能は、`ItemSend` イベントの種類によってトリガーされ、UI はありません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="7fd81-109">送信時機能に関連する制限事項の詳細については、この記事で後述する「[制限事項](#limitations)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="7fd81-110">サポートされているクライアントとプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7fd81-110">Supported clients and platforms</span></span>

<span data-ttu-id="7fd81-111">次の表に、送信時機能でサポートされているクライアントとサーバーの組み合わせを示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-111">The following table shows supported client-server combinations for the on-send feature.</span></span> <span data-ttu-id="7fd81-112">除外する組み合わせはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="7fd81-113">クライアント</span><span class="sxs-lookup"><span data-stu-id="7fd81-113">Client</span></span> | <span data-ttu-id="7fd81-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="7fd81-114">Exchange Online</span></span> | <span data-ttu-id="7fd81-115">Exchange 2016 オンプレミス</span><span class="sxs-lookup"><span data-stu-id="7fd81-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="7fd81-116">(累積更新プログラム6以降)</span><span class="sxs-lookup"><span data-stu-id="7fd81-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="7fd81-117">Exchange 2019 オンプレミス</span><span class="sxs-lookup"><span data-stu-id="7fd81-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="7fd81-118">(累積的な更新プログラム1以降)</span><span class="sxs-lookup"><span data-stu-id="7fd81-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="7fd81-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="7fd81-119">Windows:</span></span><br><span data-ttu-id="7fd81-120">バージョン 1910 (ビルド 12130.20272) 以降</span><span class="sxs-lookup"><span data-stu-id="7fd81-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="7fd81-121">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-121">Yes</span></span>|<span data-ttu-id="7fd81-122">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-122">Yes</span></span>|<span data-ttu-id="7fd81-123">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-123">Yes</span></span>|
|<span data-ttu-id="7fd81-124">Mac</span><span class="sxs-lookup"><span data-stu-id="7fd81-124">Mac:</span></span><br><span data-ttu-id="7fd81-125">ビルド16.30 以降</span><span class="sxs-lookup"><span data-stu-id="7fd81-125">build 16.30 or later</span></span>|<span data-ttu-id="7fd81-126">必要</span><span class="sxs-lookup"><span data-stu-id="7fd81-126">Yes</span></span>|<span data-ttu-id="7fd81-127">いいえ</span><span class="sxs-lookup"><span data-stu-id="7fd81-127">No</span></span>|<span data-ttu-id="7fd81-128">いいえ</span><span class="sxs-lookup"><span data-stu-id="7fd81-128">No</span></span>|
|<span data-ttu-id="7fd81-129">Web ブラウザー:</span><span class="sxs-lookup"><span data-stu-id="7fd81-129">Web browser:</span></span><br><span data-ttu-id="7fd81-130">モダン Outlook UI</span><span class="sxs-lookup"><span data-stu-id="7fd81-130">modern Outlook UI</span></span>|<span data-ttu-id="7fd81-131">あり</span><span class="sxs-lookup"><span data-stu-id="7fd81-131">Yes</span></span>|<span data-ttu-id="7fd81-132">該当なし</span><span class="sxs-lookup"><span data-stu-id="7fd81-132">Not applicable</span></span>|<span data-ttu-id="7fd81-133">該当なし</span><span class="sxs-lookup"><span data-stu-id="7fd81-133">Not applicable</span></span>|
|<span data-ttu-id="7fd81-134">Web ブラウザー:</span><span class="sxs-lookup"><span data-stu-id="7fd81-134">Web browser:</span></span><br><span data-ttu-id="7fd81-135">従来の Outlook UI</span><span class="sxs-lookup"><span data-stu-id="7fd81-135">classic Outlook UI</span></span>|<span data-ttu-id="7fd81-136">該当なし</span><span class="sxs-lookup"><span data-stu-id="7fd81-136">Not applicable</span></span>|<span data-ttu-id="7fd81-137">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-137">Yes</span></span>|<span data-ttu-id="7fd81-138">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="7fd81-139">送信時機能は、要件セット1.8 でリリースされました。詳細については、「 [現在のサーバーとクライアントのサポート](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-139">The on-send feature was released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7fd81-140">送信時機能を使用するアドインは、 [Appsource](https://appsource.microsoft.com)では許可されていません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-140">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="7fd81-141">送信時機能のしくみ</span><span class="sxs-lookup"><span data-stu-id="7fd81-141">How does the on-send feature work?</span></span>

<span data-ttu-id="7fd81-142">送信時機能を使用して、`ItemSend` 同期イベントを統合する Outlook アドインをビルドできます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-142">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="7fd81-143">このイベントは、ユーザーが [**送信**] ボタン (または、既存の会議の場合は [**変更内容を送信**] ボタン) を押していることを検出し、検証が失敗した場合はアイテムの送信をブロックするために使用できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-143">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="7fd81-144">たとえば、ユーザーがメッセージ送信イベントをトリガーすると、送信時機能を使用する Outlook アドインでは次のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-144">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="7fd81-145">電子メール メッセージの内容の読み取りと検証</span><span class="sxs-lookup"><span data-stu-id="7fd81-145">Read and validate the email message contents</span></span>
- <span data-ttu-id="7fd81-146">メッセージに件名が含まれていることの確認</span><span class="sxs-lookup"><span data-stu-id="7fd81-146">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="7fd81-147">あらかじめ定義された受信者の設定</span><span class="sxs-lookup"><span data-stu-id="7fd81-147">Set a predetermined recipient</span></span>

<span data-ttu-id="7fd81-148">検証は、送信イベントが発生したときに Outlook のクライアント側で行われ、アドインはタイムアウトするまで最大5分間あります。検証が失敗すると、アイテムの送信がブロックされ、ユーザーにアクションを実行するように求めるエラーメッセージが情報バーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-148">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

<span data-ttu-id="7fd81-149">次のスクリーンショットは、送信者に件名を追加するように通知する情報バーを示しています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-149">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![ユーザーに不足している件名を入力するように求めるエラー メッセージを示すスクリーンショット](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="7fd81-151">次のスクリーンショットは、送信者に禁止された単語が見つかったことを通知する情報バーを示しています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-151">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![送信者に禁止された単語が見つかったことを通知するエラー メッセージを示すスクリーンショット](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="7fd81-153">制限事項</span><span class="sxs-lookup"><span data-stu-id="7fd81-153">Limitations</span></span>

<span data-ttu-id="7fd81-154">現在、送信時機能には次の制限事項があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-154">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="7fd81-155">本文を呼び出す場合は、**追加-送信**機能 (プレビュー) &ndash; [。AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) on 送信ハンドラーでは、エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-155">**Append-on-send** feature (preview) &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="7fd81-156">**AppSource** &ndash; 送信時機能を使用する Outlook アドインは AppSource の検証で失敗するため、[AppSource](https://appsource.microsoft.com) に発行することはできません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-156">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="7fd81-157">送信時機能を使用するアドインは、管理者が展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-157">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="7fd81-158">**マニフェスト** &ndash; 1 つのアドインに対して 1 つの `ItemSend` イベントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-158">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="7fd81-159">マニフェストに 2 つ以上の `ItemSend` イベントがある場合、マニフェストの検証は失敗します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-159">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="7fd81-p106">**パフォーマンス** &ndash; アドインをホストする Web サーバーへの複数回のラウンドトリップは、アドインのパフォーマンスに影響する可能性があります。複数のメッセージ ベースまたは会議ベースの操作が必要なアドインを作成する場合は、パフォーマンスへの影響を考慮してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-p106">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="7fd81-162">**後で送信** (Mac のみ) &ndash; 送信時アドインがある場合、**後で送信**機能は使用できません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-162">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="7fd81-163">メールボックスの種類とモードの制限事項</span><span class="sxs-lookup"><span data-stu-id="7fd81-163">Mailbox type/mode limitations</span></span>

<span data-ttu-id="7fd81-164">送信時機能は Outlook on the web、Windows、Mac のユーザー メールボックスでのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-164">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="7fd81-165">この機能は、次のメールボックスの種類およびモードでは現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-165">The functionality is not currently supported for the following mailbox types and modes.</span></span>

- <span data-ttu-id="7fd81-166">共有メールボックス\*</span><span class="sxs-lookup"><span data-stu-id="7fd81-166">Shared mailboxes\*</span></span>
- <span data-ttu-id="7fd81-167">グループ メールボックス</span><span class="sxs-lookup"><span data-stu-id="7fd81-167">Group mailboxes</span></span>
- <span data-ttu-id="7fd81-168">オフライン モード</span><span class="sxs-lookup"><span data-stu-id="7fd81-168">Offline mode</span></span>

<span data-ttu-id="7fd81-169">送信時機能がこれらのメールボックスのシナリオに対して有効になっている場合、Outlook は送信を許可しません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-169">Outlook won't allow sending if the on-send feature is enabled for these mailbox scenarios.</span></span> <span data-ttu-id="7fd81-170">ただし、ユーザーがグループ メールボックス内のメールに返信すると、送信時アドインは実行されず、メッセージが送信されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-170">However, if a user responds to an email in a group mailbox, the on-send add-in won't run and the message will be sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7fd81-171">\* 送信時機能は、アドインが [代理人アクセスシナリオのサポートも実装](delegate-access.md)している場合は、共有メールボックスまたはフォルダーで機能します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-171">\* On-send functionality should work on shared mailboxes or folders if the add-in also [implements support for delegate access scenarios](delegate-access.md).</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="7fd81-172">複数の送信時アドイン</span><span class="sxs-lookup"><span data-stu-id="7fd81-172">Multiple on-send add-ins</span></span>

<span data-ttu-id="7fd81-173">複数の送信時アドインをインストールすると、アドインは API の `getAppManifestCall` または `getExtensibilityContext` から受信した順序で実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-173">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="7fd81-174">最初のアドインが送信を許可している場合、2 番目のアドインは最初のアドインが送信をブロックするように変更できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-174">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="7fd81-175">ただし、インストールされているすべてのアドインが送信を許可している場合、最初のアドインは再度実行されません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-175">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="7fd81-176">たとえば、アドイン 1 とアドイン 2 は両方とも送信時機能を使用しているとします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-176">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="7fd81-177">最初にアドイン 1 がインストールされ、アドイン 2 は 2 番目にインストールされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-177">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="7fd81-178">アドイン 1 は、アドインが送信を許可する条件として、Fabrikam という単語がメッセージに表示されることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-178">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="7fd81-179">ただし、アドイン 2 は Fabrikam という単語のすべての出現箇所を削除します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-179">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="7fd81-180">メッセージは、Fabrikam のすべてのインスタンスが削除されて送信されます (アドイン 1 とアドイン 2 のインストール順序のため)。</span><span class="sxs-lookup"><span data-stu-id="7fd81-180">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="7fd81-181">送信時機能を使用する Outlook アドインを展開する</span><span class="sxs-lookup"><span data-stu-id="7fd81-181">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="7fd81-182">管理者には送信時機能を使用する Outlook アドインを展開することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-182">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="7fd81-183">管理者は、送信時アドインを必ず次のようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-183">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="7fd81-184">作成項目が (電子メール、新規作成、返信、転送のために) 開かれるたびに常に存在する。</span><span class="sxs-lookup"><span data-stu-id="7fd81-184">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="7fd81-185">ユーザーが閉じたり無効にしたりできない。</span><span class="sxs-lookup"><span data-stu-id="7fd81-185">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="7fd81-186">送信時機能を使用する Outlook アドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="7fd81-186">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="7fd81-187">Outlook の送信時機能では、送信イベントの種類に対してアドインが構成されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-187">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="7fd81-188">構成するプラットフォームを選択します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-188">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="7fd81-189">Web ブラウザー - クラシック Outlook</span><span class="sxs-lookup"><span data-stu-id="7fd81-189">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="7fd81-190">送信時機能を使用する Outlook on the web (クラシック) のアドインは、*OnSendAddinsEnabled* フラグが **true** に設定された Outlook on the web メールボックス ポリシーが割り当てられているユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-190">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="7fd81-191">新しいアドインをインストールするには、次の Exchange Online PowerShell コマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-191">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="7fd81-192">リモート PowerShell を使用して Exchange Online に接続する方法については、「[Exchange Online PowerShell への接続](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-192">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="7fd81-193">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-193">Enable the on-send feature</span></span>

<span data-ttu-id="7fd81-194">既定では、送信時機能は無効になっています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-194">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="7fd81-195">管理者は、Exchange Online PowerShell コマンドレットを実行して、送信時機能を有効にできます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-195">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="7fd81-196">すべてのユーザーに対して送信時アドインを有効にするには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-196">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="7fd81-197">新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-197">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="7fd81-198">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-198">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="7fd81-199">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-199">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7fd81-200">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-200">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7fd81-201">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-201">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="7fd81-202">ユーザーのグループに対する送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-202">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="7fd81-203">ユーザーの特定のグループに対して送信時機能を有効にするための手順は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-203">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="7fd81-204">この例では、管理者は、財務担当ユーザーの環境 (財務担当ユーザーが財務部門にいる) の Outlook on the web 送信時アドイン機能のみを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-204">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="7fd81-205">グループ用の新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-205">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="7fd81-206">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています (詳細については、この記事で前述した「[メールボックスの種類の制限事項](#multiple-on-send-add-ins)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="7fd81-206">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="7fd81-207">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-207">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7fd81-208">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-208">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7fd81-209">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-209">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="7fd81-210">ポリシーが有効になるまで最大 60 分待つか、インターネット インフォメーション サービス (IIS) を再起動します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-210">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="7fd81-211">ポリシーが有効になると、グループの送信時機能が有効になります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-211">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="7fd81-212">送信時機能を無効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-212">Disable the on-send feature</span></span>

<span data-ttu-id="7fd81-213">ユーザーに対して送信時機能を無効にする、またはフラグを有効にしていない Outlook on the web のメールボックス ポリシーを割り当てるには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-213">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="7fd81-214">この例では、メールボックス ポリシーは *ContosoCorpOWAPolicy* です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-214">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="7fd81-215">**Set-OwaMailboxPolicy** コマンドレットを使用して、既存の Outlook on the web メールボックス ポリシーを構成する方法の詳細については、「[Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-215">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="7fd81-216">特定の Outlook on the web のメールボックス ポリシーが割り当てられているすべてのユーザーに対して送信時機能を無効にするには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-216">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="7fd81-217">Web ブラウザー - モダン Outlook</span><span class="sxs-lookup"><span data-stu-id="7fd81-217">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="7fd81-218">送信時機能を使用する Outlook on the web (モダン) のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-218">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7fd81-219">ただし、コンプライアンス基準を満たすためにアドインを実行する必要がある場合は、メールボックス ポリシーの *OnSendAddinsEnabled*フラグを **true** に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-219">However, if users are required to run the add-in to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="7fd81-220">新しいアドインをインストールするには、次の Exchange Online PowerShell コマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-220">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="7fd81-221">リモート PowerShell を使用して Exchange Online に接続する方法については、「[Exchange Online PowerShell への接続](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-221">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="disable-the-on-send-policy"></a><span data-ttu-id="7fd81-222">送信時ポリシーを無効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-222">Disable the on-send policy</span></span>

<span data-ttu-id="7fd81-223">既定では、送信時ポリシーは有効になっています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-223">By default, on-send policy is enabled.</span></span> <span data-ttu-id="7fd81-224">ユーザーに対して送信時ポリシーを無効にする、またはフラグを有効にしていない Outlook on the web のメールボックス ポリシーを割り当てるには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-224">To disable the on-send policy for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="7fd81-225">この例では、メールボックス ポリシーは *ContosoCorpOWAPolicy* です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-225">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="7fd81-226">**Set-OwaMailboxPolicy** コマンドレットを使用して、既存の Outlook on the web メールボックス ポリシーを構成する方法の詳細については、「[Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-226">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="7fd81-227">特定の Outlook on the web のメールボックス ポリシーが割り当てられているすべてのユーザーに対して送信時ポリシーを無効にするには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-227">To disable the on-send policy for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

#### <a name="enable-the-on-send-policy"></a><span data-ttu-id="7fd81-228">送信時ポリシーを有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-228">Enable the on-send policy</span></span>

<span data-ttu-id="7fd81-229">管理者は、Exchange Online PowerShell コマンドレットを実行して、送信時機能を有効にできます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-229">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="7fd81-230">すべてのユーザーに対して送信時アドインを有効にするには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-230">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="7fd81-231">新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-231">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="7fd81-232">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-232">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="7fd81-233">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-233">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7fd81-234">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-234">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7fd81-235">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-235">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-policy-for-a-group-of-users"></a><span data-ttu-id="7fd81-236">ユーザーのグループに対する送信時ポリシーを有効にする</span><span class="sxs-lookup"><span data-stu-id="7fd81-236">Enable the on-send policy for a group of users</span></span>

<span data-ttu-id="7fd81-237">ユーザーの特定のグループに対して送信時ポリシーを有効にするための手順は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-237">To enable the on-send policy for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="7fd81-238">この例では、管理者は、財務担当ユーザーの環境 (財務担当ユーザーが財務部門にいる) の Outlook on the web 送信時アドイン ポリシーのみを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-238">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="7fd81-239">グループ用の新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-239">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="7fd81-240">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています (詳細については、この記事で前述した「[メールボックスの種類の制限事項](#multiple-on-send-add-ins)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="7fd81-240">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="7fd81-241">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-241">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="7fd81-242">送信時ポリシーを有効にします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-242">Enable the on-send policy.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="7fd81-243">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-243">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="7fd81-244">ポリシーが有効になるまで最大 60 分待つか、インターネット インフォメーション サービス (IIS) を再起動します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-244">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="7fd81-245">ポリシーが有効になると、グループの送信時機能が適用されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-245">When the policy takes effect, the on-send feature will be enforced for the group.</span></span>

### <a name="windows"></a>[<span data-ttu-id="7fd81-246">Windows</span><span class="sxs-lookup"><span data-stu-id="7fd81-246">Windows</span></span>](#tab/windows)

<span data-ttu-id="7fd81-247">送信時機能を使用する Outlook on Windows のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-247">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7fd81-248">ただし、コンプライアンス基準を満たすためにアドインを実行する必要がある場合は、該当する各コンピュータでグループ ポリシー [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [**有効**] に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-248">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="7fd81-249">メールボックス ポリシーを設定するには、管理者が[管理用テンプレートツール](https://www.microsoft.com/download/details.aspx?id=49030)をダウンロードし、ローカル グループ ポリシー エディター (**gpedit.msc**) を実行して、最新の管理用テンプレートにアクセスします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-249">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="7fd81-250">ポリシーの内容</span><span class="sxs-lookup"><span data-stu-id="7fd81-250">What the policy does</span></span>

<span data-ttu-id="7fd81-251">コンプライアンスのために、管理者は、最新の送信時アドインを実行できるようになるまでユーザーがメッセージまたは会議アイテムを送信できないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-251">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="7fd81-252">管理者は、グループ ポリシー [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [有効] にして、すべてのアドインが Exchange から更新されるようにして、各メッセージまたは会議アイテムが予想されるルールおよび規制を送信時に満たしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-252">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="7fd81-253">ポリシーの状態</span><span class="sxs-lookup"><span data-stu-id="7fd81-253">Policy status</span></span>|<span data-ttu-id="7fd81-254">結果</span><span class="sxs-lookup"><span data-stu-id="7fd81-254">Result</span></span>|
|---|---|
|<span data-ttu-id="7fd81-255">無効</span><span class="sxs-lookup"><span data-stu-id="7fd81-255">Disabled</span></span>|<span data-ttu-id="7fd81-256">送信可能。</span><span class="sxs-lookup"><span data-stu-id="7fd81-256">Send allowed.</span></span> <span data-ttu-id="7fd81-257">アドインが Exchange からまだ更新されていない場合でも、送信時アドインを実行せずにメッセージまたは会議アイテムを送信できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-257">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="7fd81-258">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-258">Enabled</span></span>|<span data-ttu-id="7fd81-259">アドインが Exchange から更新されている場合にのみ、送信できます。それ以外の場合は、送信はブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-259">Send allowed only when the add-in has been updated from Exchange; otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="7fd81-260">送信時ポリシーを管理する</span><span class="sxs-lookup"><span data-stu-id="7fd81-260">Manage the on-send policy</span></span>

<span data-ttu-id="7fd81-261">既定では、送信時ポリシーは無効になっています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-261">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="7fd81-262">管理者は、ユーザーのグループ ポリシー設定 [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [**有効**] にすることで、送信時ポリシーを有効にできます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-262">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="7fd81-263">ユーザーのポリシーを無効にするには、管理者が [**無効**] に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-263">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="7fd81-264">このポリシー設定を管理するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="7fd81-264">To manage this policy setting, you can do the following.</span></span>

1. <span data-ttu-id="7fd81-265">最新の[管理用テンプレートツール](https://www.microsoft.com/download/details.aspx?id=49030)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-265">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="7fd81-266">ローカル グループ ポリシー エディター (**gpedit.msc**) を開きます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-266">Open the Local Group Policy editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="7fd81-267">**[ユーザーの設定] > [管理用テンプレート] > [Microsoft Outlook 2016] > [セキュリティ] > [セキュリティ センター]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-267">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="7fd81-268">[**Web 拡張機能が読み込まれない場合に送信を無効にする**] 設定を選択します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-268">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="7fd81-269">リンクを開いてポリシー設定を編集します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-269">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="7fd81-270">[**Web 拡張機能が読み込まれない場合に送信を無効にする**] ダイアログ ウィンドウで、必要に応じて [**有効**] または[**無効**] を選択し、[**OK**] または [**適用**]を選択して更新を有効にします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-270">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="7fd81-271">Mac</span><span class="sxs-lookup"><span data-stu-id="7fd81-271">Mac</span></span>](#tab/unix)

<span data-ttu-id="7fd81-272">送信時機能を使用する Outlook on Mac のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-272">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="7fd81-273">ただし、コンプライアンス基準を満たすためにアドインを実行する必要がある場合は、ユーザーの各マシンで次のメールボックス ポリシーを適用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-273">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="7fd81-274">この設定またはキーは、CFPreferences と互換性があります。つまり、Jamf Pro などの Mac のエンタープライズ管理ソフトウェアを使用して設定することができます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-274">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

|||
|:---|:---|
|<span data-ttu-id="7fd81-275">**ドメイン**</span><span class="sxs-lookup"><span data-stu-id="7fd81-275">**Domain**</span></span>|<span data-ttu-id="7fd81-276">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="7fd81-276">com.microsoft.outlook</span></span>|
|<span data-ttu-id="7fd81-277">**キー**</span><span class="sxs-lookup"><span data-stu-id="7fd81-277">**Key**</span></span>|<span data-ttu-id="7fd81-278">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="7fd81-278">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="7fd81-279">**DataType**</span><span class="sxs-lookup"><span data-stu-id="7fd81-279">**DataType**</span></span>|<span data-ttu-id="7fd81-280">Boolean</span><span class="sxs-lookup"><span data-stu-id="7fd81-280">Boolean</span></span>|
|<span data-ttu-id="7fd81-281">**指定可能な値**</span><span class="sxs-lookup"><span data-stu-id="7fd81-281">**Possible values**</span></span>|<span data-ttu-id="7fd81-282">false (既定)</span><span class="sxs-lookup"><span data-stu-id="7fd81-282">false (default)</span></span><br><span data-ttu-id="7fd81-283">true</span><span class="sxs-lookup"><span data-stu-id="7fd81-283">true</span></span>|
|<span data-ttu-id="7fd81-284">**可用性**</span><span class="sxs-lookup"><span data-stu-id="7fd81-284">**Availability**</span></span>|<span data-ttu-id="7fd81-285">16.27</span><span class="sxs-lookup"><span data-stu-id="7fd81-285">16.27</span></span>|
|<span data-ttu-id="7fd81-286">**コメント**</span><span class="sxs-lookup"><span data-stu-id="7fd81-286">**Comments**</span></span>|<span data-ttu-id="7fd81-287">このキーは、送信時メールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-287">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="7fd81-288">設定内容</span><span class="sxs-lookup"><span data-stu-id="7fd81-288">What the setting does</span></span>

<span data-ttu-id="7fd81-289">コンプライアンスのために、管理者は、最新の送信時アドインを実行できるようになるまでユーザーがメッセージまたは会議アイテムを送信できないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-289">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="7fd81-290">管理者は、キー **OnSendAddinsWaitForLoad** を有効にして、すべてのアドインが Exchange から更新されるようにして、各メッセージまたは会議アイテムが予想されるルールおよび規制を送信時に満たしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-290">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="7fd81-291">キーの状態</span><span class="sxs-lookup"><span data-stu-id="7fd81-291">Key's state</span></span>|<span data-ttu-id="7fd81-292">結果</span><span class="sxs-lookup"><span data-stu-id="7fd81-292">Result</span></span>|
|---|---|
|<span data-ttu-id="7fd81-293">false</span><span class="sxs-lookup"><span data-stu-id="7fd81-293">false</span></span>|<span data-ttu-id="7fd81-294">送信可能。</span><span class="sxs-lookup"><span data-stu-id="7fd81-294">Send allowed.</span></span> <span data-ttu-id="7fd81-295">アドインが Exchange からまだ更新されていない場合でも、送信時アドインを実行せずにメッセージまたは会議アイテムを送信できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-295">Message or meeting item can be sent without running the on-send add-in, even if the add-in has not been updated from Exchange yet.</span></span>|
|<span data-ttu-id="7fd81-296">true</span><span class="sxs-lookup"><span data-stu-id="7fd81-296">true</span></span>|<span data-ttu-id="7fd81-297">アドインが Exchange から更新されている場合にのみ、送信できます。それ以外の場合は、送信はブロックされ、[**Send**] ボタンは無効です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-297">Send allowed only when add-ins have been updated from Exchange; otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="7fd81-298">送信時機能のシナリオ</span><span class="sxs-lookup"><span data-stu-id="7fd81-298">On-send feature scenarios</span></span>

<span data-ttu-id="7fd81-299">送信時機能を使用するアドインのサポートされているシナリオとサポートされていないシナリオは、次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-299">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="7fd81-300">ユーザー メールボックスで送信時アドイン機能が有効になっているが、アドインはインストールされていない</span><span class="sxs-lookup"><span data-stu-id="7fd81-300">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="7fd81-301">このシナリオでは、ユーザーはアドインを実行せずにメッセージおよび会議アイテムを送信することができます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-301">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="7fd81-302">ユーザー メールボックスで送信時アドイン機能が有効になっており、送信時機能をサポートするアドインがインストールされ、有効になっている</span><span class="sxs-lookup"><span data-stu-id="7fd81-302">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="7fd81-303">アドインは送信イベント中に実行され、ユーザーによる送信を許可またはブロックします。</span><span class="sxs-lookup"><span data-stu-id="7fd81-303">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="7fd81-304">メールボックス 1 がメールボックス 2 への完全なアクセス許可を持つ、メールボックスの委任</span><span class="sxs-lookup"><span data-stu-id="7fd81-304">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="7fd81-305">Web ブラウザー (クラシック Outlook)</span><span class="sxs-lookup"><span data-stu-id="7fd81-305">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="7fd81-306">シナリオ</span><span class="sxs-lookup"><span data-stu-id="7fd81-306">Scenario</span></span>|<span data-ttu-id="7fd81-307">メールボックス 1 の送信時機能</span><span class="sxs-lookup"><span data-stu-id="7fd81-307">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="7fd81-308">メールボックス 2 の送信時機能</span><span class="sxs-lookup"><span data-stu-id="7fd81-308">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="7fd81-309">Outlook web のセッション (クラシック)</span><span class="sxs-lookup"><span data-stu-id="7fd81-309">Outlook web session (classic)</span></span>|<span data-ttu-id="7fd81-310">結果</span><span class="sxs-lookup"><span data-stu-id="7fd81-310">Result</span></span>|<span data-ttu-id="7fd81-311">サポートの有無</span><span class="sxs-lookup"><span data-stu-id="7fd81-311">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="7fd81-312">1 </span><span class="sxs-lookup"><span data-stu-id="7fd81-312">1</span></span>|<span data-ttu-id="7fd81-313">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-313">Enabled</span></span>|<span data-ttu-id="7fd81-314">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-314">Enabled</span></span>|<span data-ttu-id="7fd81-315">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="7fd81-315">New session</span></span>|<span data-ttu-id="7fd81-316">メールボックス 1 は、メールボックス 2 からのメッセージまたは会議アイテムを送信できません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-316">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="7fd81-p132">現在サポートされていません。回避策として、シナリオ 3 を使用します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-p132">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="7fd81-319">2 </span><span class="sxs-lookup"><span data-stu-id="7fd81-319">2</span></span>|<span data-ttu-id="7fd81-320">無効</span><span class="sxs-lookup"><span data-stu-id="7fd81-320">Disabled</span></span>|<span data-ttu-id="7fd81-321">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-321">Enabled</span></span>|<span data-ttu-id="7fd81-322">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="7fd81-322">New session</span></span>|<span data-ttu-id="7fd81-323">メールボックス 1 は、メールボックス 2 からのメッセージまたは会議アイテムを送信できません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-323">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="7fd81-p133">現在サポートされていません。回避策として、シナリオ 3 を使用します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-p133">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="7fd81-326">3 </span><span class="sxs-lookup"><span data-stu-id="7fd81-326">3</span></span>|<span data-ttu-id="7fd81-327">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-327">Enabled</span></span>|<span data-ttu-id="7fd81-328">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-328">Enabled</span></span>|<span data-ttu-id="7fd81-329">同じセッション</span><span class="sxs-lookup"><span data-stu-id="7fd81-329">Same session</span></span>|<span data-ttu-id="7fd81-330">メールボックス 1 に割り当てられている送信時アドインが送信時に実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-330">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="7fd81-331">サポートされています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-331">Supported.</span></span>|
|<span data-ttu-id="7fd81-332">4 </span><span class="sxs-lookup"><span data-stu-id="7fd81-332">4</span></span>|<span data-ttu-id="7fd81-333">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-333">Enabled</span></span>|<span data-ttu-id="7fd81-334">無効</span><span class="sxs-lookup"><span data-stu-id="7fd81-334">Disabled</span></span>|<span data-ttu-id="7fd81-335">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="7fd81-335">New session</span></span>|<span data-ttu-id="7fd81-336">送信時アドインは実行されません。メッセージまたは会議アイテムは送信されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-336">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="7fd81-337">サポートされています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-337">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="7fd81-338">Web ブラウザー (モダン Outlook)、Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="7fd81-338">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="7fd81-339">強制的に送信するには、管理者は両方のメールボックスでポリシーが有効になっていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-339">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="7fd81-340">アドインで代理人アクセスをサポートする方法については、「[Outlook アドインでの代理人アクセスのシナリオを有効にする](delegate-access.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-340">To learn how to support delegate access in an add-in, see [Enable delegate access scenarios in an Outlook add-in](delegate-access.md).</span></span>

### <a name="group-1-is-a-modern-group-mailbox-and-user-mailbox-1-is-a-member-of-group-1"></a><span data-ttu-id="7fd81-341">グループ 1 がモダン グループ メールボックスであり、ユーザー メールボックス 1 がグループ 1 のメンバーである</span><span class="sxs-lookup"><span data-stu-id="7fd81-341">Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1</span></span>

<br/>

|<span data-ttu-id="7fd81-342">シナリオ</span><span class="sxs-lookup"><span data-stu-id="7fd81-342">Scenario</span></span>|<span data-ttu-id="7fd81-343">メールボックス 1 の送信時ポリシー</span><span class="sxs-lookup"><span data-stu-id="7fd81-343">Mailbox 1 on-send policy</span></span>|<span data-ttu-id="7fd81-344">送信時アドインが有効かどうか</span><span class="sxs-lookup"><span data-stu-id="7fd81-344">On-send add-ins enabled?</span></span>|<span data-ttu-id="7fd81-345">メールボックス 1 のアクション</span><span class="sxs-lookup"><span data-stu-id="7fd81-345">Mailbox 1 action</span></span>|<span data-ttu-id="7fd81-346">結果</span><span class="sxs-lookup"><span data-stu-id="7fd81-346">Result</span></span>|<span data-ttu-id="7fd81-347">サポートの有無</span><span class="sxs-lookup"><span data-stu-id="7fd81-347">Supported?</span></span>|
|:------------|:-------------------------|:-------------------|:---------|:----------|:-------------|
|<span data-ttu-id="7fd81-348">1 </span><span class="sxs-lookup"><span data-stu-id="7fd81-348">1</span></span>|<span data-ttu-id="7fd81-349">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-349">Enabled</span></span>|<span data-ttu-id="7fd81-350">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-350">Yes</span></span>|<span data-ttu-id="7fd81-351">メールボックス 1 はグループ 1 への新しいメッセージまたは会議を作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-351">Mailbox 1 composes new message or meeting to Group 1.</span></span>|<span data-ttu-id="7fd81-352">送信中に送信時アドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-352">On-send add-ins run during send.</span></span>|<span data-ttu-id="7fd81-353">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-353">Yes</span></span>|
|<span data-ttu-id="7fd81-354">2 </span><span class="sxs-lookup"><span data-stu-id="7fd81-354">2</span></span>|<span data-ttu-id="7fd81-355">有効</span><span class="sxs-lookup"><span data-stu-id="7fd81-355">Enabled</span></span>|<span data-ttu-id="7fd81-356">はい</span><span class="sxs-lookup"><span data-stu-id="7fd81-356">Yes</span></span>|<span data-ttu-id="7fd81-357">メールボックス 1 は、Outlook on the web のグループ 1 のグループ ウィンドウ内でグループ 1 への新しいメッセージまたは会議を作成します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-357">Mailbox 1 composes a new message or meeting to Group 1 within Group 1's group window in Outlook on the web.</span></span>|<span data-ttu-id="7fd81-358">送信中に送信時アドインは実行されません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-358">On-send add-ins do not run during send.</span></span>|<span data-ttu-id="7fd81-359">現在サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-359">Not currently supported.</span></span> <span data-ttu-id="7fd81-360">回避策として、シナリオ 1 を使用します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-360">As a workaround, use scenario 1.</span></span>|

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="7fd81-361">ユーザー メールボックスで送信時アドイン機能/ポリシーが有効になっており、送信時機能をサポートするアドインがインストールされ、有効であり、オフライン モードが有効になっている</span><span class="sxs-lookup"><span data-stu-id="7fd81-361">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="7fd81-362">送信時アドインは、ユーザー、アドイン バックエンド、および Exchange のオンライン状態に従って実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-362">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="7fd81-363">ユーザーの状態</span><span class="sxs-lookup"><span data-stu-id="7fd81-363">User's state</span></span>

<span data-ttu-id="7fd81-364">ユーザーがオンラインの場合、送信中に送信時アドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-364">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="7fd81-365">ユーザーがオフラインの場合、送信中に送信時アドインは実行されず、メッセージまたは会議アイテムは送信されません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-365">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="7fd81-366">アドイン バックエンドの状態</span><span class="sxs-lookup"><span data-stu-id="7fd81-366">Add-in backend's state</span></span>

<span data-ttu-id="7fd81-367">送信時アドインは、バックエンドがオンラインで接続可能な場合に実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-367">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="7fd81-368">バックエンドがオフラインの場合、送信は無効です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-368">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="7fd81-369">Exchange の状態</span><span class="sxs-lookup"><span data-stu-id="7fd81-369">Exchange's state</span></span>

<span data-ttu-id="7fd81-370">Exchange サーバーがオンラインでアクセスできる場合、送信中に送信時アドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-370">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="7fd81-371">送信時アドインが Exchange に接続できない場合、および該当するポリシーまたはコマンドレットが有効になっている場合、送信は無効です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-371">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="7fd81-372">オフライン状態の Mac では [**送信**] ボタン (または、既存の会議の場合は [**変更内容を送信**] ボタン) が無効になっており、ユーザーがオフラインの場合、組織が送信を許可していないという通知が表示されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-372">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="7fd81-373">送信中に、ユーザーがアイテムを編集できるようになります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-373">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="7fd81-374">送信中のアドインはアイテムを処理していますが、ユーザーは不適切なテキストや添付ファイルなどを追加することでアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-374">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="7fd81-375">アドインが送信時に処理している間にユーザーがアイテムを編集できないようにするには、ダイアログを使用して回避策を実装します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-375">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="7fd81-376">送信ハンドラーの場合:</span><span class="sxs-lookup"><span data-stu-id="7fd81-376">In your on-send handler:</span></span>

1. <span data-ttu-id="7fd81-377">[Displaydialogasync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-)を呼び出して、マウスクリックとキーボード操作が無効になるようにダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-377">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="7fd81-378">Web 上の Outlook でこの動作を取得するには、呼び出しのパラメーターで [Displayiniframe プロパティ](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) をに設定する必要があり `true` `options` `displayDialogAsync` ます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-378">To get this behavior in Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="7fd81-379">アイテムの処理を実装します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-379">Implement processing of the item.</span></span>
1. <span data-ttu-id="7fd81-380">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-380">Close the dialog.</span></span> <span data-ttu-id="7fd81-381">また、ユーザーがダイアログを閉じた場合の処理についても処理します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-381">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="7fd81-382">コード例</span><span class="sxs-lookup"><span data-stu-id="7fd81-382">Code examples</span></span>

<span data-ttu-id="7fd81-383">次のコード例は、単純な送信時アドインを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-383">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="7fd81-384">これらの例を基にしたコード サンプルをダウンロードするには、「[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-384">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="7fd81-385">送信イベントでダイアログを使用する場合は、イベントを完了する前に必ずダイアログを閉じてください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-385">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="7fd81-386">マニフェスト、バージョンのオーバーライド、イベント</span><span class="sxs-lookup"><span data-stu-id="7fd81-386">Manifest, version override, and event</span></span>

<span data-ttu-id="7fd81-387">[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) コード サンプルには、2 つのマニフェストが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-387">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="7fd81-388">`Contoso Message Body Checker.xml` &ndash; 制限された単語または機密情報についてメッセージの本文を送信時に確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-388">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="7fd81-389">`Contoso Subject and CC Checker.xml` &ndash; CC 行に受信者を追加し、送信時にメッセージに件名が含まれていることを確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-389">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="7fd81-390">`Contoso Message Body Checker.xml` マニフェスト ファイルには、`ItemSend` イベントで呼び出す必要がある関数ファイルと関数名を含めます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-390">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="7fd81-391">操作は同期的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-391">The operation runs synchronously.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case, the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

> [!IMPORTANT]
> <span data-ttu-id="7fd81-392">送信時アドインを開発するために Visual Studio 2019 を使用している場合は、次のような検証警告が表示されることがあります。 "これは、無効な xsi: type ' http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events ' です" です。これを回避するには、 [この警告についてのブログ](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)で GitHub gist として提供されている新しいバージョンの MailAppVersionOverridesV1_1 が必要になります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-392">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="7fd81-393">`Contoso Subject and CC Checker.xml` マニフェスト ファイルの場合、次の例では、メッセージ送信イベントで呼び出す関数ファイルと関数名を示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-393">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
                <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
        </DesktopFormFactor>
    </Host>
</Hosts>
```

<br/>

<span data-ttu-id="7fd81-394">送信時 API には `VersionOverrides v1_1` が必要です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-394">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="7fd81-395">マニフェストに `VersionOverrides` ノードを追加する方法を次に示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-395">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="7fd81-396">詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-396">For more information, see the following:</span></span>
> - [<span data-ttu-id="7fd81-397">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="7fd81-397">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="7fd81-398">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="7fd81-398">Office Add-ins XML manifest</span></span>](../overview/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="7fd81-399">`Event` オブジェクト、`item` オブジェクトと、`body.getAsync` メソッド、`body.setAsync` メソッドを理解する</span><span class="sxs-lookup"><span data-stu-id="7fd81-399">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="7fd81-400">現在選択されているメッセージまたは会議アイテム (この例では、新しく作成されたメッセージ) にアクセスするには、`Office.context.mailbox.item` 名前空間を使用します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-400">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="7fd81-401">`ItemSend` イベントは、送信時機能によってマニフェストで指定された関数に自動的に渡されます &mdash; この例では `validateBody` 関数です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-401">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

```js
var mailboxItem;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}
```

<span data-ttu-id="7fd81-402">`validateBody` 関数は、指定した形式 (HTML) の現在の本文を取得し、コールバック メソッドでのアクセスにコードが必要とする `ItemSend` イベント オブジェクトを渡します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-402">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="7fd81-403">`getAsync` メソッドに加え、`Body` オブジェクトは本文を指定したテキストに置き換えるために使用できる `setAsync` メソッドも提供します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-403">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="7fd81-404">詳細については、「[Event オブジェクト](/javascript/api/office/office.addincommands.event)」と「[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-404">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="7fd81-405">`NotificationMessages` オブジェクトと `event.completed` メソッド</span><span class="sxs-lookup"><span data-stu-id="7fd81-405">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="7fd81-406">`checkBodyOnlyOnSendCallBack` 関数は、正規表現を使用して、禁止された単語がメッセージの本文に含まれているかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-406">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="7fd81-407">制限されている単語の配列に対する一致が検出された場合、電子メールの送信をブロックし、情報バーを使用して送信者に通知します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-407">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="7fd81-408">これを実行するには、`Item` オブジェクトの `notificationMessages` プロパティを使用して、`NotificationMessages` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-408">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="7fd81-409">その後、次の例に示すように、`addAsync` メソッドを呼び出して通知をアイテムに追加します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-409">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

```js
// Determine whether the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allow sending.
// <param name="asyncResult">ItemSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
    var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
    var wordExpression = listOfBlockedWords.join('|');

    // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
    // i to perform case-insensitive search.
    var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
    var checkBody = regexCheck.test(asyncResult.value);

    if (checkBody) {
        mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
        // Block send.
        asyncResult.asyncContext.completed({ allowEvent: false });
    }

    // Allow send.
    asyncResult.asyncContext.completed({ allowEvent: true });
}
```

<span data-ttu-id="7fd81-410">`addAsync` メソッドのパラメーターは、次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-410">The following are the parameters for the `addAsync` method:</span></span>

- <span data-ttu-id="7fd81-411">`NoSend` &ndash; 通知メッセージを参照するための開発者が指定したキーである文字列。</span><span class="sxs-lookup"><span data-stu-id="7fd81-411">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="7fd81-412">これを使用して後でこのメッセージを変更できます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-412">You can use it to modify this message later.</span></span> <span data-ttu-id="7fd81-413">キーの長さは32文字以内にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="7fd81-413">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="7fd81-414">`type` &ndash; JSON オブジェクト パラメーターのプロパティの 1 つ。</span><span class="sxs-lookup"><span data-stu-id="7fd81-414">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="7fd81-415">メッセージの種類を表します。種類は [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) 列挙型の値に対応しています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-415">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="7fd81-416">使用可能な値は、進行状況のインジケーター、情報メッセージ、エラー メッセージです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-416">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="7fd81-417">この例では、`type` はエラー メッセージです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-417">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="7fd81-418">`message` &ndash; JSON オブジェクト パラメーターのプロパティの 1 つ。</span><span class="sxs-lookup"><span data-stu-id="7fd81-418">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="7fd81-419">この例では、`message` は通知メッセージのテキストです。</span><span class="sxs-lookup"><span data-stu-id="7fd81-419">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="7fd81-420">アドインが送信操作によってトリガーされた `ItemSend` イベントの処理を完了したことを通知するには、`event.completed({allowEvent:Boolean})` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-420">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="7fd81-421">`allowEvent` プロパティは Boolean です。</span><span class="sxs-lookup"><span data-stu-id="7fd81-421">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="7fd81-422">`true` に設定されている場合、送信が許可されます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-422">If set to `true`, send is allowed.</span></span> <span data-ttu-id="7fd81-423">`false` に設定されている場合、電子メール メッセージの送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="7fd81-423">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="7fd81-424">詳細については、「[notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)」と「[completed](/javascript/api/office/office.addincommands.event)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-424">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="7fd81-425">`replaceAsync` メソッド、`removeAsync` メソッド、および`getAllAsync`メソッド</span><span class="sxs-lookup"><span data-stu-id="7fd81-425">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="7fd81-426">`addAsync` メソッドに加え、`NotificationMessages` オブジェクトは本文を指定したテキストに置き換えるために使用できる `replaceAsync`、`removeAsync`、および `getAllAsync` の各メソッドも提供します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-426">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="7fd81-427">このコード サンプルでは、これらのメソッドは使用されません。</span><span class="sxs-lookup"><span data-stu-id="7fd81-427">These methods are not used in this code sample.</span></span>  <span data-ttu-id="7fd81-428">詳細については、「[NotificationMessages](/javascript/api/outlook/office.NotificationMessages)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7fd81-428">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="7fd81-429">件名および CC のチェッカー コード</span><span class="sxs-lookup"><span data-stu-id="7fd81-429">Subject and CC checker code</span></span>

<span data-ttu-id="7fd81-430">次のコード例では、CC 行に受信者を追加し、送信時にメッセージに件名が含まれていることを確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-430">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="7fd81-431">この例では、送信時機能を使用して、電子メールの送信を許可または禁止します。</span><span class="sxs-lookup"><span data-stu-id="7fd81-431">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

```js
// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">ItemSend event is automatically passed by on-send code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
    shouldChangeSubjectOnSend(event);
}

// Determine whether the subject should be changed. If it is already changed, allow send. Otherwise change it.
// <param name="event">ItemSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
    mailboxItem.subject.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Determine whether a string is blank, null, or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                if (!checkSubject) {
                    subjectOnSendChange(subject, asyncResult.asyncContext);
                    //console.log(checkSubject);
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            }
        });
}

// Add a CC to the email. In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">ItemSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });
}

// Determine whether the subject should be changed. If it is already changed, allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">ItemSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                // Block send.
                asyncResult.asyncContext.completed({ allowEvent: false });
            }
            else {
                // Allow send.
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        });
}
```

<span data-ttu-id="7fd81-p153">CC 行に受信者を追加して、送信時にメッセージに件名が含まれていることを確認する方法、および使用可能な API を表示する方法の詳細については、「[Outlook-Add-in-On-Send サンプル](https://github.com/OfficeDev/Outlook-Add-in-On-Send)」を参照してください。コードには詳細なコメントが付けられています。</span><span class="sxs-lookup"><span data-stu-id="7fd81-p153">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="7fd81-434">関連項目</span><span class="sxs-lookup"><span data-stu-id="7fd81-434">See also</span></span>

- [<span data-ttu-id="7fd81-435">Outlook アドインのアーキテクチャと機能の概要</span><span class="sxs-lookup"><span data-stu-id="7fd81-435">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="7fd81-436">アドイン コマンド デモの Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="7fd81-436">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)
