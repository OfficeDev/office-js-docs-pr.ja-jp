---
title: Outlook アドインの送信時機能
description: アイテムを処理する方法、またはユーザーが特定のアクションを実行しないようにする方法を提供し、送信時にアドインが特定のプロパティを設定できるようにします。
ms.date: 06/16/2021
localization_priority: Normal
ms.openlocfilehash: 80047f4c8056bafa62d467f1e69dd334d168486a
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348476"
---
# <a name="on-send-feature-for-outlook-add-ins"></a><span data-ttu-id="36862-103">Outlook アドインの送信時機能</span><span class="sxs-lookup"><span data-stu-id="36862-103">On-send feature for Outlook add-ins</span></span>

<span data-ttu-id="36862-p101">Outlook アドインの送信時機能は、メッセージまたは会議アイテムを処理する方法、またはユーザーが特定のアクションを実行できないようにする方法を提供し、送信時にアドインが特定のプロパティを設定できるようにします。たとえば、送信時機能を使用すると次のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="36862-p101">The on-send feature for Outlook add-ins provides a way to handle a message or meeting item, or block users from certain actions, and allows an add-in to set certain properties on send. For example, you can use the on-send feature to:</span></span>

- <span data-ttu-id="36862-106">ユーザーが機密情報を送信したり、件名を空白にしたままにしないようにする。</span><span class="sxs-lookup"><span data-stu-id="36862-106">Prevent a user from sending sensitive information or leaving the subject line blank.</span></span>  
- <span data-ttu-id="36862-107">特定の受信者をメッセージの CC 行または会議の任意の受信者行に追加する。</span><span class="sxs-lookup"><span data-stu-id="36862-107">Add a specific recipient to the CC line in messages, or to the optional recipients line in meetings.</span></span>

<span data-ttu-id="36862-108">送信時の機能は、`ItemSend` イベントの種類によってトリガーされ、UI はありません。</span><span class="sxs-lookup"><span data-stu-id="36862-108">The on-send feature is triggered by the `ItemSend` event type and is UI-less.</span></span>

<span data-ttu-id="36862-109">送信時機能に関連する制限事項の詳細については、この記事で後述する「[制限事項](#limitations)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-109">For information about limitations related to the on-send feature, see [Limitations](#limitations) later in this article.</span></span>

## <a name="supported-clients-and-platforms"></a><span data-ttu-id="36862-110">サポートされているクライアントとプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="36862-110">Supported clients and platforms</span></span>

<span data-ttu-id="36862-111">次の表に、必要な最小累積更新プログラム (該当する場合) を含む、送信時機能でサポートされるクライアントとサーバーの組み合わせを示します。</span><span class="sxs-lookup"><span data-stu-id="36862-111">The following table shows supported client-server combinations for the on-send feature, including the minimum required Cumulative Update where applicable.</span></span> <span data-ttu-id="36862-112">除外された組み合わせはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="36862-112">Excluded combinations are not supported.</span></span>

| <span data-ttu-id="36862-113">クライアント</span><span class="sxs-lookup"><span data-stu-id="36862-113">Client</span></span> | <span data-ttu-id="36862-114">Exchange Online</span><span class="sxs-lookup"><span data-stu-id="36862-114">Exchange Online</span></span> | <span data-ttu-id="36862-115">Exchange 2016 オンプレミス</span><span class="sxs-lookup"><span data-stu-id="36862-115">Exchange 2016 on-premises</span></span><br><span data-ttu-id="36862-116">(累積的な更新プログラム 6 以降)</span><span class="sxs-lookup"><span data-stu-id="36862-116">(Cumulative Update 6 or later)</span></span> | <span data-ttu-id="36862-117">Exchange 2019 オンプレミス</span><span class="sxs-lookup"><span data-stu-id="36862-117">Exchange 2019 on-premises</span></span><br><span data-ttu-id="36862-118">(累積的な更新プログラム 1 以降)</span><span class="sxs-lookup"><span data-stu-id="36862-118">(Cumulative Update 1 or later)</span></span> |
|---|:---:|:---:|:---:|
|<span data-ttu-id="36862-119">Windows:</span><span class="sxs-lookup"><span data-stu-id="36862-119">Windows:</span></span><br><span data-ttu-id="36862-120">バージョン 1910 (ビルド 12130.20272) 以降</span><span class="sxs-lookup"><span data-stu-id="36862-120">version 1910 (build 12130.20272) or later</span></span>|<span data-ttu-id="36862-121">はい</span><span class="sxs-lookup"><span data-stu-id="36862-121">Yes</span></span>|<span data-ttu-id="36862-122">はい</span><span class="sxs-lookup"><span data-stu-id="36862-122">Yes</span></span>|<span data-ttu-id="36862-123">はい</span><span class="sxs-lookup"><span data-stu-id="36862-123">Yes</span></span>|
|<span data-ttu-id="36862-124">Mac:</span><span class="sxs-lookup"><span data-stu-id="36862-124">Mac:</span></span><br><span data-ttu-id="36862-125">ビルド 16.47 以降</span><span class="sxs-lookup"><span data-stu-id="36862-125">build 16.47 or later</span></span>|<span data-ttu-id="36862-126">はい</span><span class="sxs-lookup"><span data-stu-id="36862-126">Yes</span></span>|<span data-ttu-id="36862-127">はい</span><span class="sxs-lookup"><span data-stu-id="36862-127">Yes</span></span>|<span data-ttu-id="36862-128">はい</span><span class="sxs-lookup"><span data-stu-id="36862-128">Yes</span></span>|
|<span data-ttu-id="36862-129">Web ブラウザー:</span><span class="sxs-lookup"><span data-stu-id="36862-129">Web browser:</span></span><br><span data-ttu-id="36862-130">モダン Outlook UI</span><span class="sxs-lookup"><span data-stu-id="36862-130">modern Outlook UI</span></span>|<span data-ttu-id="36862-131">あり</span><span class="sxs-lookup"><span data-stu-id="36862-131">Yes</span></span>|<span data-ttu-id="36862-132">該当なし</span><span class="sxs-lookup"><span data-stu-id="36862-132">Not applicable</span></span>|<span data-ttu-id="36862-133">該当なし</span><span class="sxs-lookup"><span data-stu-id="36862-133">Not applicable</span></span>|
|<span data-ttu-id="36862-134">Web ブラウザー:</span><span class="sxs-lookup"><span data-stu-id="36862-134">Web browser:</span></span><br><span data-ttu-id="36862-135">クラシック Outlook UI</span><span class="sxs-lookup"><span data-stu-id="36862-135">classic Outlook UI</span></span>|<span data-ttu-id="36862-136">該当なし</span><span class="sxs-lookup"><span data-stu-id="36862-136">Not applicable</span></span>|<span data-ttu-id="36862-137">はい</span><span class="sxs-lookup"><span data-stu-id="36862-137">Yes</span></span>|<span data-ttu-id="36862-138">はい</span><span class="sxs-lookup"><span data-stu-id="36862-138">Yes</span></span>|

> [!NOTE]
> <span data-ttu-id="36862-139">オン送信機能は、要件セット 1.8 で正式にリリースされました (詳細については、現在のサーバーと [クライアントのサポートを](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) 参照してください)。</span><span class="sxs-lookup"><span data-stu-id="36862-139">The on-send feature was officially released in requirement set 1.8 (see [current server and client support](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details).</span></span> <span data-ttu-id="36862-140">ただし、機能のサポート マトリックスは要件セットのスーパーセットです。</span><span class="sxs-lookup"><span data-stu-id="36862-140">However, note that the feature's support matrix is a superset of the requirement set's.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="36862-141">送信時機能を使用するアドインは [、AppSource では許可されません](https://appsource.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="36862-141">Add-ins that use the on-send feature aren't allowed in [AppSource](https://appsource.microsoft.com).</span></span>

## <a name="how-does-the-on-send-feature-work"></a><span data-ttu-id="36862-142">送信時機能のしくみ</span><span class="sxs-lookup"><span data-stu-id="36862-142">How does the on-send feature work?</span></span>

<span data-ttu-id="36862-143">送信時機能を使用して、`ItemSend` 同期イベントを統合する Outlook アドインをビルドできます。</span><span class="sxs-lookup"><span data-stu-id="36862-143">You can use the on-send feature to build an Outlook add-in that integrates the `ItemSend` synchronous event.</span></span> <span data-ttu-id="36862-144">このイベントは、ユーザーが [**送信**] ボタン (または、既存の会議の場合は [**変更内容を送信**] ボタン) を押していることを検出し、検証が失敗した場合はアイテムの送信をブロックするために使用できます。</span><span class="sxs-lookup"><span data-stu-id="36862-144">This event detects that the user is pressing the **Send** button (or the **Send Update** button for existing meetings) and can be used to block the item from sending if the validation fails.</span></span> <span data-ttu-id="36862-145">たとえば、ユーザーがメッセージ送信イベントをトリガーすると、送信時機能を使用する Outlook アドインでは次のことが可能です。</span><span class="sxs-lookup"><span data-stu-id="36862-145">For example, when a user triggers a message send event, an Outlook add-in that uses the on-send feature can:</span></span>

- <span data-ttu-id="36862-146">電子メール メッセージの内容の読み取りと検証</span><span class="sxs-lookup"><span data-stu-id="36862-146">Read and validate the email message contents</span></span>
- <span data-ttu-id="36862-147">メッセージに件名が含まれていることの確認</span><span class="sxs-lookup"><span data-stu-id="36862-147">Verify that the message includes a subject line</span></span>
- <span data-ttu-id="36862-148">あらかじめ定義された受信者の設定</span><span class="sxs-lookup"><span data-stu-id="36862-148">Set a predetermined recipient</span></span>

<span data-ttu-id="36862-149">検証は、送信イベントがOutlookのクライアント側で行われ、アドインがタイム アウトする前に最大 5 分かかります。検証に失敗すると、アイテムの送信がブロックされ、ユーザーにアクションを実行するように求めるエラー メッセージが情報バーに表示されます。</span><span class="sxs-lookup"><span data-stu-id="36862-149">Validation is done on the client side in Outlook when the send event is triggered, and the add-in has up to 5 minutes before it times out. If validation fails, the sending of the item is blocked, and an error message is displayed in an information bar that prompts the user to take action.</span></span>

> [!NOTE]
> <span data-ttu-id="36862-150">Outlook on the web では、Outlook ブラウザー タブ内で構成されているメッセージで送信時の機能がトリガーされると、検証や他の処理を完了するために、アイテムが独自のブラウザー ウィンドウまたはタブにポップアウトされます。</span><span class="sxs-lookup"><span data-stu-id="36862-150">In Outlook on the web, when the on-send feature is triggered in a message being composed within the Outlook browser tab, the item is popped out to its own browser window or tab in order to complete validation and other processing.</span></span>

<span data-ttu-id="36862-151">次のスクリーンショットは、送信者に件名を追加するように通知する情報バーを示しています。</span><span class="sxs-lookup"><span data-stu-id="36862-151">The following screenshot shows an information bar that notifies the sender to add a subject.</span></span>

<br/>

![不足している件名行を入力するようにユーザーに求めるエラー メッセージを示すスクリーンショット。](../images/block-on-send-subject-cc-inforbar.png)

<br/>

<br/>

<span data-ttu-id="36862-153">次のスクリーンショットは、送信者に禁止された単語が見つかったことを通知する情報バーを示しています。</span><span class="sxs-lookup"><span data-stu-id="36862-153">The following screenshot shows an information bar that notifies the sender that blocked words were found.</span></span>

<br/>

![ブロックされた単語が見つかったことをユーザーに伝えるエラー メッセージを示すスクリーンショット。](../images/block-on-send-body.png)

## <a name="limitations"></a><span data-ttu-id="36862-155">制限事項</span><span class="sxs-lookup"><span data-stu-id="36862-155">Limitations</span></span>

<span data-ttu-id="36862-156">現在、送信時機能には次の制限事項があります。</span><span class="sxs-lookup"><span data-stu-id="36862-156">The on-send feature currently has the following limitations.</span></span>

- <span data-ttu-id="36862-157">**Append-on-send フィーチャー** &ndash; 本文を呼び出す [場合。送信時ハンドラーの AppendOnSendAsync、](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) エラーが返されます。</span><span class="sxs-lookup"><span data-stu-id="36862-157">**Append-on-send** feature &ndash; If you call [body.AppendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-) in the on-send handler, an error is returned.</span></span>
- <span data-ttu-id="36862-158">**AppSource** &ndash; 送信時機能を使用する Outlook アドインは AppSource の検証で失敗するため、[AppSource](https://appsource.microsoft.com) に発行することはできません。</span><span class="sxs-lookup"><span data-stu-id="36862-158">**AppSource** &ndash; You can't publish Outlook add-ins that use the on-send feature to [AppSource](https://appsource.microsoft.com) as they will fail AppSource validation.</span></span> <span data-ttu-id="36862-159">送信時機能を使用するアドインは、管理者が展開する必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-159">Add-ins that use the on-send feature should be deployed by administrators.</span></span>
- <span data-ttu-id="36862-160">**マニフェスト** &ndash; 1 つのアドインに対して 1 つの `ItemSend` イベントのみがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="36862-160">**Manifest** &ndash; Only one `ItemSend` event is supported per add-in.</span></span> <span data-ttu-id="36862-161">マニフェストに 2 つ以上の `ItemSend` イベントがある場合、マニフェストの検証は失敗します。</span><span class="sxs-lookup"><span data-stu-id="36862-161">If you have two or more `ItemSend` events in a manifest, the manifest will fail validation.</span></span>
- <span data-ttu-id="36862-p107">**パフォーマンス** &ndash; アドインをホストする Web サーバーへの複数回のラウンドトリップは、アドインのパフォーマンスに影響する可能性があります。複数のメッセージ ベースまたは会議ベースの操作が必要なアドインを作成する場合は、パフォーマンスへの影響を考慮してください。</span><span class="sxs-lookup"><span data-stu-id="36862-p107">**Performance** &ndash; Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple message- or meeting-based operations.</span></span>
- <span data-ttu-id="36862-164">**後で送信** (Mac のみ) &ndash; 送信時アドインがある場合、**後で送信** 機能は使用できません。</span><span class="sxs-lookup"><span data-stu-id="36862-164">**Send Later** (Mac only) &ndash; If there are on-send add-ins, the **Send Later** feature will be unavailable.</span></span>

<span data-ttu-id="36862-165">また、イベントの完了後にアイテムを閉じると自動的に行われますので、送信時のイベント ハンドラーを呼び出 `item.close()` すのは推奨されません。</span><span class="sxs-lookup"><span data-stu-id="36862-165">Also, it's not recommended that you call `item.close()` in the on-send event handler as closing the item should happen automatically after the event is completed.</span></span>

### <a name="mailbox-typemode-limitations"></a><span data-ttu-id="36862-166">メールボックスの種類とモードの制限事項</span><span class="sxs-lookup"><span data-stu-id="36862-166">Mailbox type/mode limitations</span></span>

<span data-ttu-id="36862-167">送信時機能は Outlook on the web、Windows、Mac のユーザー メールボックスでのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="36862-167">On-send functionality is only supported for user mailboxes in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="36862-168">Outlook アドインの概要ページの [アドインで使用できるメールボックス アイテム][](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)セクションで、アドインがアクティブ化されない状況に加えて、機能は現在オフライン モードではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="36862-168">In addition to situations where add-ins don't activate as noted in the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page, the functionality is not currently supported for offline mode.</span></span>

<span data-ttu-id="36862-169">Outlookがサポートされていないメールボックス シナリオでオン送信機能が有効になっている場合、送信は許可されません。</span><span class="sxs-lookup"><span data-stu-id="36862-169">Outlook won't allow sending if the on-send feature is enabled for unsupported mailbox scenarios.</span></span> <span data-ttu-id="36862-170">ただし、Outlookアドインがアクティブ化しない場合、送信時アドインは実行されません。メッセージが送信されます。</span><span class="sxs-lookup"><span data-stu-id="36862-170">However, in cases where Outlook add-ins don't activate, the on-send add-in won't run and the message will be sent.</span></span>

## <a name="multiple-on-send-add-ins"></a><span data-ttu-id="36862-171">複数の送信時アドイン</span><span class="sxs-lookup"><span data-stu-id="36862-171">Multiple on-send add-ins</span></span>

<span data-ttu-id="36862-172">複数の送信時アドインをインストールすると、アドインは API の `getAppManifestCall` または `getExtensibilityContext` から受信した順序で実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-172">If multiple on-send add-ins are installed, the add-ins will run in the order in which they are received from APIs `getAppManifestCall` or `getExtensibilityContext`.</span></span> <span data-ttu-id="36862-173">最初のアドインが送信を許可している場合、2 番目のアドインは最初のアドインが送信をブロックするように変更できます。</span><span class="sxs-lookup"><span data-stu-id="36862-173">If the first add-in allows sending, the second add-in can change something that would make the first one block sending.</span></span> <span data-ttu-id="36862-174">ただし、インストールされているすべてのアドインが送信を許可している場合、最初のアドインは再度実行されません。</span><span class="sxs-lookup"><span data-stu-id="36862-174">However, the first add-in won't run again if all installed add-ins have allowed sending.</span></span>

<span data-ttu-id="36862-175">たとえば、アドイン 1 とアドイン 2 は両方とも送信時機能を使用しているとします。</span><span class="sxs-lookup"><span data-stu-id="36862-175">For example, Add-in1 and Add-in2 both use the on-send feature.</span></span> <span data-ttu-id="36862-176">最初にアドイン 1 がインストールされ、アドイン 2 は 2 番目にインストールされます。</span><span class="sxs-lookup"><span data-stu-id="36862-176">Add-in1 is installed first, and Add-in2 is installed second.</span></span> <span data-ttu-id="36862-177">アドイン 1 は、アドインが送信を許可する条件として、Fabrikam という単語がメッセージに表示されることを確認します。</span><span class="sxs-lookup"><span data-stu-id="36862-177">Add-in1 verifies that the word Fabrikam appears in the message as a condition for the add-in to allow send.</span></span>  <span data-ttu-id="36862-178">ただし、アドイン 2 は Fabrikam という単語のすべての出現箇所を削除します。</span><span class="sxs-lookup"><span data-stu-id="36862-178">However, Add-in2 removes any occurrences of the word Fabrikam.</span></span> <span data-ttu-id="36862-179">メッセージは、Fabrikam のすべてのインスタンスが削除されて送信されます (アドイン 1 とアドイン 2 のインストール順序のため)。</span><span class="sxs-lookup"><span data-stu-id="36862-179">The message will send with all instances of Fabrikam removed (due to the order of installation of Add-in1 and Add-in2).</span></span>

## <a name="deploy-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="36862-180">送信時機能を使用する Outlook アドインを展開する</span><span class="sxs-lookup"><span data-stu-id="36862-180">Deploy Outlook add-ins that use on-send</span></span>

<span data-ttu-id="36862-181">管理者には送信時機能を使用する Outlook アドインを展開することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="36862-181">We recommend that administrators deploy Outlook add-ins that use the on-send feature.</span></span> <span data-ttu-id="36862-182">管理者は、送信時アドインを必ず次のようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-182">Administrators have to ensure that the on-send add-in:</span></span>

- <span data-ttu-id="36862-183">作成項目が (電子メール、新規作成、返信、転送のために) 開かれるたびに常に存在する。</span><span class="sxs-lookup"><span data-stu-id="36862-183">Is always present any time a compose item is opened (for email: new, reply, or forward).</span></span>
- <span data-ttu-id="36862-184">ユーザーが閉じたり無効にしたりできない。</span><span class="sxs-lookup"><span data-stu-id="36862-184">Can't be closed or disabled by the user.</span></span>

## <a name="install-outlook-add-ins-that-use-on-send"></a><span data-ttu-id="36862-185">送信時機能を使用する Outlook アドインをインストールする</span><span class="sxs-lookup"><span data-stu-id="36862-185">Install Outlook add-ins that use on-send</span></span>

<span data-ttu-id="36862-186">Outlook の送信時機能では、送信イベントの種類に対してアドインが構成されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-186">The on-send feature in Outlook requires that add-ins are configured for the send event types.</span></span> <span data-ttu-id="36862-187">構成するプラットフォームを選択します。</span><span class="sxs-lookup"><span data-stu-id="36862-187">Select the platform you'd like to configure.</span></span>

### <a name="web-browser---classic-outlook"></a>[<span data-ttu-id="36862-188">Web ブラウザー - クラシック Outlook</span><span class="sxs-lookup"><span data-stu-id="36862-188">Web browser - classic Outlook</span></span>](#tab/classic)

<span data-ttu-id="36862-189">送信時機能を使用する Outlook on the web (クラシック) のアドインは、*OnSendAddinsEnabled* フラグが **true** に設定された Outlook on the web メールボックス ポリシーが割り当てられているユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-189">Add-ins for Outlook on the web (classic) that use the on-send feature will run for users who are assigned an Outlook on the web mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.</span></span>

<span data-ttu-id="36862-190">新しいアドインをインストールするには、次の Exchange Online PowerShell コマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-190">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="36862-191">リモート PowerShell を使用して Exchange Online に接続する方法については、「[Exchange Online PowerShell への接続](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-191">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-feature"></a><span data-ttu-id="36862-192">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="36862-192">Enable the on-send feature</span></span>

<span data-ttu-id="36862-193">既定では、送信時機能は無効になっています。</span><span class="sxs-lookup"><span data-stu-id="36862-193">By default, on-send functionality is disabled.</span></span> <span data-ttu-id="36862-194">管理者は、Exchange Online PowerShell コマンドレットを実行して、送信時機能を有効にできます。</span><span class="sxs-lookup"><span data-stu-id="36862-194">Administrators can enable on-send by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="36862-195">すべてのユーザーに対して送信時アドインを有効にするには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="36862-195">To enable on-send add-ins for all users:</span></span>

1. <span data-ttu-id="36862-196">新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="36862-196">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="36862-197">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="36862-197">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="36862-198">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-198">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="36862-199">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="36862-199">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="36862-200">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="36862-200">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="enable-the-on-send-feature-for-a-group-of-users"></a><span data-ttu-id="36862-201">ユーザーのグループに対する送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="36862-201">Enable the on-send feature for a group of users</span></span>

<span data-ttu-id="36862-202">ユーザーの特定のグループに対して送信時機能を有効にするための手順は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="36862-202">To enable the on-send feature for a specific group of users the steps are as follows.</span></span>  <span data-ttu-id="36862-203">この例では、管理者は、財務担当ユーザーの環境 (財務担当ユーザーが財務部門にいる) の Outlook on the web 送信時アドイン機能のみを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-203">In this example, an administrator only wants to enable an Outlook on the web on-send add-in feature in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="36862-204">グループ用の新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="36862-204">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="36862-205">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています (詳細については、この記事で前述した「[メールボックスの種類の制限事項](#multiple-on-send-add-ins)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="36862-205">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="36862-206">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-206">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="36862-207">送信時機能を有効にする</span><span class="sxs-lookup"><span data-stu-id="36862-207">Enable the on-send feature.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="36862-208">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="36862-208">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="36862-209">ポリシーが有効になるまで最大 60 分待つか、インターネット インフォメーション サービス (IIS) を再起動します。</span><span class="sxs-lookup"><span data-stu-id="36862-209">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="36862-210">ポリシーが有効になると、グループの送信時機能が有効になります。</span><span class="sxs-lookup"><span data-stu-id="36862-210">When the policy takes effect, the on-send feature will be enabled for the group.</span></span>

#### <a name="disable-the-on-send-feature"></a><span data-ttu-id="36862-211">送信時機能を無効にする</span><span class="sxs-lookup"><span data-stu-id="36862-211">Disable the on-send feature</span></span>

<span data-ttu-id="36862-212">ユーザーに対して送信時機能を無効にする、またはフラグを有効にしていない Outlook on the web のメールボックス ポリシーを割り当てるには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-212">To disable the on-send feature for a user or assign an Outlook on the web mailbox policy that does not have the flag enabled, run the following cmdlets.</span></span> <span data-ttu-id="36862-213">この例では、メールボックス ポリシーは *ContosoCorpOWAPolicy* です。</span><span class="sxs-lookup"><span data-stu-id="36862-213">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="36862-214">**Set-OwaMailboxPolicy** コマンドレットを使用して、既存の Outlook on the web メールボックス ポリシーを構成する方法の詳細については、「[Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-214">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="36862-215">特定の Outlook on the web のメールボックス ポリシーが割り当てられているすべてのユーザーに対して送信時機能を無効にするには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-215">To disable the on-send feature for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="36862-216">Web ブラウザー - モダン Outlook</span><span class="sxs-lookup"><span data-stu-id="36862-216">Web browser - modern Outlook</span></span>](#tab/modern)

<span data-ttu-id="36862-217">送信時機能を使用する Outlook on the web (モダン) のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-217">Add-ins for Outlook on the web (modern) that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="36862-218">ただし、ユーザーがコンプライアンス基準を満たすために送信時アドインを実行する必要がある場合は、メールボックス ポリシーに *OnSendAddinsEnabled* フラグを設定して、アドインの送信時にアイテムの編集が許可されない必要があります。 `true`</span><span class="sxs-lookup"><span data-stu-id="36862-218">However, if users are required to run on-send add-ins to meet compliance standards, then the mailbox policy must have the *OnSendAddinsEnabled* flag set to `true` so that editing the item is not allowed while the add-ins are processing on send.</span></span>

<span data-ttu-id="36862-219">新しいアドインをインストールするには、次の Exchange Online PowerShell コマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-219">To install a new add-in, run the following Exchange Online PowerShell cmdlets.</span></span>

```powershell
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0
```

```powershell
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> [!NOTE]
> <span data-ttu-id="36862-220">リモート PowerShell を使用して Exchange Online に接続する方法については、「[Exchange Online PowerShell への接続](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-220">To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell).</span></span>

#### <a name="enable-the-on-send-flag"></a><span data-ttu-id="36862-221">送信時フラグを有効にする</span><span class="sxs-lookup"><span data-stu-id="36862-221">Enable the on-send flag</span></span>

<span data-ttu-id="36862-222">管理者は、PowerShell コマンドレットを実行して、Exchange Onlineコンプライアンスを適用できます。</span><span class="sxs-lookup"><span data-stu-id="36862-222">Administrators can enforce on-send compliance by running Exchange Online PowerShell cmdlets.</span></span>

<span data-ttu-id="36862-223">すべてのユーザーに対して、オン送信アドインの処理中に編集を禁止するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="36862-223">For all users, to disallow editing while on-send add-ins are processing:</span></span>

1. <span data-ttu-id="36862-224">新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="36862-224">Create a new Outlook on the web mailbox policy.</span></span>

   ```powershell
    New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

    > [!NOTE]
    > <span data-ttu-id="36862-225">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="36862-225">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types.</span></span> <span data-ttu-id="36862-226">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-226">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="36862-227">送信時にコンプライアンスを適用します。</span><span class="sxs-lookup"><span data-stu-id="36862-227">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="36862-228">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="36862-228">Assign the policy to users.</span></span>

   ```powershell
    Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
   ```

#### <a name="turn-on-the-on-send-flag-for-a-group-of-users"></a><span data-ttu-id="36862-229">ユーザーのグループの送信時フラグをオンにする</span><span class="sxs-lookup"><span data-stu-id="36862-229">Turn on the on-send flag for a group of users</span></span>

<span data-ttu-id="36862-230">特定のユーザー グループに対して送信時のコンプライアンスを適用するには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-230">To enforce on-send compliance for a specific group of users, the steps are as follows.</span></span> <span data-ttu-id="36862-231">この例では、管理者は、財務担当ユーザーの環境 (財務担当ユーザーが財務部門にいる) の Outlook on the web 送信時アドイン ポリシーのみを有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-231">In this example, an administrator only wants to enable an Outlook on the web on-send add-in policy in an environment for Finance users (where the Finance users are in the Finance Department).</span></span>

1. <span data-ttu-id="36862-232">グループ用の新しい Outlook on the web のメールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="36862-232">Create a new Outlook on the web mailbox policy for the group.</span></span>

   ```powershell
    New-OWAMailboxPolicy FinanceOWAPolicy
   ```

   > [!NOTE]
   > <span data-ttu-id="36862-233">管理者は既存のポリシーを使用できますが、送信時機能は特定のメールボックスの種類でのみサポートされています (詳細については、この記事で前述した「[メールボックスの種類の制限事項](#multiple-on-send-add-ins)」を参照してください)。</span><span class="sxs-lookup"><span data-stu-id="36862-233">Administrators can use an existing policy, but on-send functionality is only supported on certain mailbox types (see [Mailbox type limitations](#multiple-on-send-add-ins) earlier in this article for more information).</span></span> <span data-ttu-id="36862-234">サポートされていないメールボックスは、Outlook on the web では既定で送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-234">Unsupported mailboxes will be blocked from sending by default in Outlook on the web.</span></span>

2. <span data-ttu-id="36862-235">送信時にコンプライアンスを適用します。</span><span class="sxs-lookup"><span data-stu-id="36862-235">Enforce compliance on send.</span></span>

   ```powershell
    Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
   ```

3. <span data-ttu-id="36862-236">ポリシーをユーザーに割り当てます。</span><span class="sxs-lookup"><span data-stu-id="36862-236">Assign the policy to users.</span></span>

   ```powershell
    $targetUsers = Get-Group 'Finance'|select -ExpandProperty members
    $targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
   ```

> [!NOTE]
> <span data-ttu-id="36862-237">ポリシーが有効になるまで最大 60 分待つか、インターネット インフォメーション サービス (IIS) を再起動します。</span><span class="sxs-lookup"><span data-stu-id="36862-237">Wait up to 60 minutes for the policy to take effect, or restart Internet Information Services (IIS).</span></span> <span data-ttu-id="36862-238">ポリシーが有効な場合、グループに対して送信時のコンプライアンスが適用されます。</span><span class="sxs-lookup"><span data-stu-id="36862-238">When the policy takes effect, on-send compliance will be enforced for the group.</span></span>

#### <a name="turn-off-the-on-send-flag"></a><span data-ttu-id="36862-239">送信時フラグをオフにする</span><span class="sxs-lookup"><span data-stu-id="36862-239">Turn off the on-send flag</span></span>

<span data-ttu-id="36862-240">ユーザーの送信時コンプライアンスの適用を無効にするには、次のコマンドレットを実行してフラグを有効にしていない Outlook on the web メールボックス ポリシーを割り当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-240">To turn off on-send compliance enforcement for a user, assign an Outlook on the web mailbox policy that does not have the flag enabled by running the following cmdlets.</span></span> <span data-ttu-id="36862-241">この例では、メールボックス ポリシーは *ContosoCorpOWAPolicy* です。</span><span class="sxs-lookup"><span data-stu-id="36862-241">In this example, the mailbox policy is *ContosoCorpOWAPolicy*.</span></span>

```powershell
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

> [!NOTE]
> <span data-ttu-id="36862-242">**Set-OwaMailboxPolicy** コマンドレットを使用して、既存の Outlook on the web メールボックス ポリシーを構成する方法の詳細については、「[Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-242">For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](/powershell/module/exchange/client-access/Set-OwaMailboxPolicy).</span></span>

<span data-ttu-id="36862-243">特定のメールボックス ポリシーが割り当てられているすべてのユーザーに対するオンOutlook on the web適用を無効にするには、次のコマンドレットを実行します。</span><span class="sxs-lookup"><span data-stu-id="36862-243">To turn off on-send compliance enforcement for all users that have a specific Outlook on the web mailbox policy assigned, run the following cmdlets.</span></span>

```powershell
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

### <a name="windows"></a>[<span data-ttu-id="36862-244">Windows</span><span class="sxs-lookup"><span data-stu-id="36862-244">Windows</span></span>](#tab/windows)

<span data-ttu-id="36862-245">送信時機能を使用する Outlook on Windows のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-245">Add-ins for Outlook on Windows that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="36862-246">ただし、コンプライアンス基準を満たすためにアドインを実行する必要がある場合は、該当する各コンピュータでグループ ポリシー [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [**有効**] に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-246">However, if users are required to run the add-in to meet compliance standards, then the group policy **Disable send when web extensions can't load** must be set to **Enabled** on each applicable machine.</span></span>

<span data-ttu-id="36862-247">メールボックス ポリシーを設定するには、管理者は管理用 [](https://www.microsoft.com/download/details.aspx?id=49030)テンプレート ツールをダウンロードし、ローカル グループ ポリシー エディター **gpedit.msc** を実行して最新の管理テンプレートにアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="36862-247">To set mailbox policies, administrators can download the [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030) then access the latest administrative templates by running the Local Group Policy Editor, **gpedit.msc**.</span></span>

#### <a name="what-the-policy-does"></a><span data-ttu-id="36862-248">ポリシーの内容</span><span class="sxs-lookup"><span data-stu-id="36862-248">What the policy does</span></span>

<span data-ttu-id="36862-249">コンプライアンスのために、管理者は、最新の送信時アドインを実行できるようになるまでユーザーがメッセージまたは会議アイテムを送信できないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-249">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-in is available to run.</span></span> <span data-ttu-id="36862-250">管理者は、グループ ポリシー [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [有効] にして、すべてのアドインが Exchange から更新されるようにして、各メッセージまたは会議アイテムが予想されるルールおよび規制を送信時に満たしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="36862-250">Administrators must enable the group policy **Disable send when web extensions can't load** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="36862-251">ポリシーの状態</span><span class="sxs-lookup"><span data-stu-id="36862-251">Policy status</span></span>|<span data-ttu-id="36862-252">結果</span><span class="sxs-lookup"><span data-stu-id="36862-252">Result</span></span>|
|---|---|
|<span data-ttu-id="36862-253">無効</span><span class="sxs-lookup"><span data-stu-id="36862-253">Disabled</span></span>|<span data-ttu-id="36862-254">送信時アドインの現在ダウンロードされているマニフェスト (必ずしも最新バージョンではない) は、送信されるメッセージまたは会議アイテムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-254">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="36862-255">これは既定の状態/動作です。</span><span class="sxs-lookup"><span data-stu-id="36862-255">This is the default status/behavior.</span></span>|
|<span data-ttu-id="36862-256">Enabled</span><span class="sxs-lookup"><span data-stu-id="36862-256">Enabled</span></span>|<span data-ttu-id="36862-257">送信時アドインの最新のマニフェストが Exchange からダウンロードされると、送信されるメッセージまたは会議アイテムに対してアドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-257">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="36862-258">それ以外の場合、送信はブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-258">Otherwise, send is blocked.</span></span>|

#### <a name="manage-the-on-send-policy"></a><span data-ttu-id="36862-259">送信時ポリシーを管理する</span><span class="sxs-lookup"><span data-stu-id="36862-259">Manage the on-send policy</span></span>

<span data-ttu-id="36862-260">既定では、送信時ポリシーは無効になっています。</span><span class="sxs-lookup"><span data-stu-id="36862-260">By default, the on-send policy is disabled.</span></span> <span data-ttu-id="36862-261">管理者は、ユーザーのグループ ポリシー設定 [**Web 拡張機能が読み込まれない場合に送信を無効にする**] を [**有効**] にすることで、送信時ポリシーを有効にできます。</span><span class="sxs-lookup"><span data-stu-id="36862-261">Administrators can enable the on-send policy by ensuring the user's group policy setting **Disable send when web extensions can't load** is set to **Enabled**.</span></span> <span data-ttu-id="36862-262">ユーザーのポリシーを無効にするには、管理者が [**無効**] に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-262">To disable the policy for a user, the administrator should set it to **Disabled**.</span></span> <span data-ttu-id="36862-263">このポリシー設定を管理するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="36862-263">To manage this policy setting, you can do the following:</span></span>

1. <span data-ttu-id="36862-264">最新の[管理用テンプレートツール](https://www.microsoft.com/download/details.aspx?id=49030)をダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="36862-264">Download the latest [Administrative Templates tool](https://www.microsoft.com/download/details.aspx?id=49030).</span></span>
1. <span data-ttu-id="36862-265">ローカル グループ ポリシー エディター **(gpedit.msc) を開きます**。</span><span class="sxs-lookup"><span data-stu-id="36862-265">Open the Local Group Policy Editor (**gpedit.msc**).</span></span>
1. <span data-ttu-id="36862-266">**[ユーザーの設定] > [管理用テンプレート] > [Microsoft Outlook 2016] > [セキュリティ] > [セキュリティ センター]** の順に移動します。</span><span class="sxs-lookup"><span data-stu-id="36862-266">Navigate to **User Configuration > Administrative Templates  > Microsoft Outlook 2016 > Security > Trust Center**.</span></span>
1. <span data-ttu-id="36862-267">[**Web 拡張機能が読み込まれない場合に送信を無効にする**] 設定を選択します。</span><span class="sxs-lookup"><span data-stu-id="36862-267">Select the **Disable send when web extensions can't load** setting.</span></span>
1. <span data-ttu-id="36862-268">リンクを開いてポリシー設定を編集します。</span><span class="sxs-lookup"><span data-stu-id="36862-268">Open the link to edit policy setting.</span></span>
1. <span data-ttu-id="36862-269">[**Web 拡張機能が読み込まれない場合に送信を無効にする**] ダイアログ ウィンドウで、必要に応じて [**有効**] または[**無効**] を選択し、[**OK**] または [**適用**]を選択して更新を有効にします。</span><span class="sxs-lookup"><span data-stu-id="36862-269">In the **Disable send when web extensions can't load** dialog window, select **Enabled** or **Disabled** as appropriate then select **OK** or **Apply** to put the update into effect.</span></span>

### <a name="mac"></a>[<span data-ttu-id="36862-270">Mac</span><span class="sxs-lookup"><span data-stu-id="36862-270">Mac</span></span>](#tab/unix)

<span data-ttu-id="36862-271">送信時機能を使用する Outlook on Mac のアドインは、インストールされているすべてのユーザーに対して実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-271">Add-ins for Outlook on Mac that use the on-send feature should run for any users who have them installed.</span></span> <span data-ttu-id="36862-272">ただし、コンプライアンス基準を満たすためにアドインを実行する必要がある場合は、ユーザーの各マシンで次のメールボックス ポリシーを適用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-272">However, if users are required to run the add-in to meet compliance standards, then the following mailbox setting must be applied on each user's machine.</span></span> <span data-ttu-id="36862-273">この設定またはキーは、CFPreferences と互換性があります。つまり、Jamf Pro などの Mac のエンタープライズ管理ソフトウェアを使用して設定することができます。</span><span class="sxs-lookup"><span data-stu-id="36862-273">This setting or key is CFPreference-compatible, which means that it can be set by using enterprise management software for Mac, such as Jamf Pro.</span></span>

||<span data-ttu-id="36862-274">値</span><span class="sxs-lookup"><span data-stu-id="36862-274">Value</span></span>|
|:---|:---|
|<span data-ttu-id="36862-275">**ドメイン**</span><span class="sxs-lookup"><span data-stu-id="36862-275">**Domain**</span></span>|<span data-ttu-id="36862-276">com.microsoft.outlook</span><span class="sxs-lookup"><span data-stu-id="36862-276">com.microsoft.outlook</span></span>|
|<span data-ttu-id="36862-277">**キー**</span><span class="sxs-lookup"><span data-stu-id="36862-277">**Key**</span></span>|<span data-ttu-id="36862-278">OnSendAddinsWaitForLoad</span><span class="sxs-lookup"><span data-stu-id="36862-278">OnSendAddinsWaitForLoad</span></span>|
|<span data-ttu-id="36862-279">**DataType**</span><span class="sxs-lookup"><span data-stu-id="36862-279">**DataType**</span></span>|<span data-ttu-id="36862-280">Boolean</span><span class="sxs-lookup"><span data-stu-id="36862-280">Boolean</span></span>|
|<span data-ttu-id="36862-281">**指定可能な値**</span><span class="sxs-lookup"><span data-stu-id="36862-281">**Possible values**</span></span>|<span data-ttu-id="36862-282">false (既定)</span><span class="sxs-lookup"><span data-stu-id="36862-282">false (default)</span></span><br><span data-ttu-id="36862-283">true</span><span class="sxs-lookup"><span data-stu-id="36862-283">true</span></span>|
|<span data-ttu-id="36862-284">**可用性**</span><span class="sxs-lookup"><span data-stu-id="36862-284">**Availability**</span></span>|<span data-ttu-id="36862-285">16.27</span><span class="sxs-lookup"><span data-stu-id="36862-285">16.27</span></span>|
|<span data-ttu-id="36862-286">**コメント**</span><span class="sxs-lookup"><span data-stu-id="36862-286">**Comments**</span></span>|<span data-ttu-id="36862-287">このキーは、送信時メールボックス ポリシーを作成します。</span><span class="sxs-lookup"><span data-stu-id="36862-287">This key creates an onSendMailbox policy.</span></span>|

#### <a name="what-the-setting-does"></a><span data-ttu-id="36862-288">設定内容</span><span class="sxs-lookup"><span data-stu-id="36862-288">What the setting does</span></span>

<span data-ttu-id="36862-289">コンプライアンスのために、管理者は、最新の送信時アドインを実行できるようになるまでユーザーがメッセージまたは会議アイテムを送信できないようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-289">For compliance reasons, administrators may need to ensure that users cannot send message or meeting items until the latest on-send add-ins are available to run.</span></span> <span data-ttu-id="36862-290">管理者は、キー **OnSendAddinsWaitForLoad** を有効にして、すべてのアドインが Exchange から更新されるようにして、各メッセージまたは会議アイテムが予想されるルールおよび規制を送信時に満たしていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="36862-290">Admins must enable the key **OnSendAddinsWaitForLoad** so that all add-ins are updated from Exchange and available to verify each message or meeting item meets expected rules and regulations on send.</span></span>

|<span data-ttu-id="36862-291">キーの状態</span><span class="sxs-lookup"><span data-stu-id="36862-291">Key's state</span></span>|<span data-ttu-id="36862-292">結果</span><span class="sxs-lookup"><span data-stu-id="36862-292">Result</span></span>|
|---|---|
|<span data-ttu-id="36862-293">false</span><span class="sxs-lookup"><span data-stu-id="36862-293">false</span></span>|<span data-ttu-id="36862-294">送信時アドインの現在ダウンロードされているマニフェスト (必ずしも最新バージョンではない) は、送信されるメッセージまたは会議アイテムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-294">The currently downloaded manifests of the on-send add-ins (not necessarily the latest versions) run on message or meeting items being sent.</span></span> <span data-ttu-id="36862-295">これは既定の状態/動作です。</span><span class="sxs-lookup"><span data-stu-id="36862-295">This is the default state/behavior.</span></span>|
|<span data-ttu-id="36862-296">true</span><span class="sxs-lookup"><span data-stu-id="36862-296">true</span></span>|<span data-ttu-id="36862-297">送信時アドインの最新のマニフェストが Exchange からダウンロードされると、送信されるメッセージまたは会議アイテムに対してアドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-297">After the latest manifests of the on-send add-ins are downloaded from Exchange, the add-ins are run on message or meeting items being sent.</span></span> <span data-ttu-id="36862-298">それ以外の場合は、送信がブロックされ、[送信] **ボタン** が無効になります。</span><span class="sxs-lookup"><span data-stu-id="36862-298">Otherwise, send is blocked and the **Send** button is disabled.</span></span>|

---

## <a name="on-send-feature-scenarios"></a><span data-ttu-id="36862-299">送信時機能のシナリオ</span><span class="sxs-lookup"><span data-stu-id="36862-299">On-send feature scenarios</span></span>

<span data-ttu-id="36862-300">送信時機能を使用するアドインのサポートされているシナリオとサポートされていないシナリオは、次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="36862-300">The following are the supported and unsupported scenarios for add-ins that use the on-send feature.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-but-no-add-ins-are-installed"></a><span data-ttu-id="36862-301">ユーザー メールボックスで送信時アドイン機能が有効になっているが、アドインはインストールされていない</span><span class="sxs-lookup"><span data-stu-id="36862-301">User mailbox has the on-send add-in feature enabled but no add-ins are installed</span></span>

<span data-ttu-id="36862-302">このシナリオでは、ユーザーはアドインを実行せずにメッセージおよび会議アイテムを送信することができます。</span><span class="sxs-lookup"><span data-stu-id="36862-302">In this scenario the user will be able to send message and meeting items without any add-ins executing.</span></span>

### <a name="user-mailbox-has-the-on-send-add-in-feature-enabled-and-add-ins-that-supports-on-send-are-installed-and-enabled"></a><span data-ttu-id="36862-303">ユーザー メールボックスで送信時アドイン機能が有効になっており、送信時機能をサポートするアドインがインストールされ、有効になっている</span><span class="sxs-lookup"><span data-stu-id="36862-303">User mailbox has the on-send add-in feature enabled and add-ins that supports on-send are installed and enabled</span></span>

<span data-ttu-id="36862-304">アドインは送信イベント中に実行され、ユーザーによる送信を許可またはブロックします。</span><span class="sxs-lookup"><span data-stu-id="36862-304">Add-ins will run during the send event, which will then either allow or block the user from sending.</span></span>

### <a name="mailbox-delegation-where-mailbox-1-has-full-access-permissions-to-mailbox-2"></a><span data-ttu-id="36862-305">メールボックス 1 がメールボックス 2 への完全なアクセス許可を持つ、メールボックスの委任</span><span class="sxs-lookup"><span data-stu-id="36862-305">Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2</span></span>

#### <a name="web-browser-classic-outlook"></a><span data-ttu-id="36862-306">Web ブラウザー (クラシック Outlook)</span><span class="sxs-lookup"><span data-stu-id="36862-306">Web browser (classic Outlook)</span></span>

|<span data-ttu-id="36862-307">シナリオ</span><span class="sxs-lookup"><span data-stu-id="36862-307">Scenario</span></span>|<span data-ttu-id="36862-308">メールボックス 1 の送信時機能</span><span class="sxs-lookup"><span data-stu-id="36862-308">Mailbox 1 on-send feature</span></span>|<span data-ttu-id="36862-309">メールボックス 2 の送信時機能</span><span class="sxs-lookup"><span data-stu-id="36862-309">Mailbox 2 on-send feature</span></span>|<span data-ttu-id="36862-310">Outlook web のセッション (クラシック)</span><span class="sxs-lookup"><span data-stu-id="36862-310">Outlook web session (classic)</span></span>|<span data-ttu-id="36862-311">結果</span><span class="sxs-lookup"><span data-stu-id="36862-311">Result</span></span>|<span data-ttu-id="36862-312">サポートの有無</span><span class="sxs-lookup"><span data-stu-id="36862-312">Supported?</span></span>|
|:------------|:------------|:--------------------------|:---------|:-------------|:-------------|
|<span data-ttu-id="36862-313">1</span><span class="sxs-lookup"><span data-stu-id="36862-313">1</span></span>|<span data-ttu-id="36862-314">有効</span><span class="sxs-lookup"><span data-stu-id="36862-314">Enabled</span></span>|<span data-ttu-id="36862-315">有効</span><span class="sxs-lookup"><span data-stu-id="36862-315">Enabled</span></span>|<span data-ttu-id="36862-316">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="36862-316">New session</span></span>|<span data-ttu-id="36862-317">メールボックス 1 は、メールボックス 2 からのメッセージまたは会議アイテムを送信できません。</span><span class="sxs-lookup"><span data-stu-id="36862-317">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="36862-p135">現在サポートされていません。回避策として、シナリオ 3 を使用します。</span><span class="sxs-lookup"><span data-stu-id="36862-p135">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="36862-320">2</span><span class="sxs-lookup"><span data-stu-id="36862-320">2</span></span>|<span data-ttu-id="36862-321">無効</span><span class="sxs-lookup"><span data-stu-id="36862-321">Disabled</span></span>|<span data-ttu-id="36862-322">有効</span><span class="sxs-lookup"><span data-stu-id="36862-322">Enabled</span></span>|<span data-ttu-id="36862-323">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="36862-323">New session</span></span>|<span data-ttu-id="36862-324">メールボックス 1 は、メールボックス 2 からのメッセージまたは会議アイテムを送信できません。</span><span class="sxs-lookup"><span data-stu-id="36862-324">Mailbox 1 cannot send a message or meeting item from mailbox 2.</span></span>|<span data-ttu-id="36862-p136">現在サポートされていません。回避策として、シナリオ 3 を使用します。</span><span class="sxs-lookup"><span data-stu-id="36862-p136">Not currently supported. As a workaround, use scenario 3.</span></span>|
|<span data-ttu-id="36862-327">3</span><span class="sxs-lookup"><span data-stu-id="36862-327">3</span></span>|<span data-ttu-id="36862-328">有効</span><span class="sxs-lookup"><span data-stu-id="36862-328">Enabled</span></span>|<span data-ttu-id="36862-329">有効</span><span class="sxs-lookup"><span data-stu-id="36862-329">Enabled</span></span>|<span data-ttu-id="36862-330">同じセッション</span><span class="sxs-lookup"><span data-stu-id="36862-330">Same session</span></span>|<span data-ttu-id="36862-331">メールボックス 1 に割り当てられている送信時アドインが送信時に実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-331">On-send add-ins assigned to mailbox 1 run on-send.</span></span>|<span data-ttu-id="36862-332">サポートされています。</span><span class="sxs-lookup"><span data-stu-id="36862-332">Supported.</span></span>|
|<span data-ttu-id="36862-333">4 </span><span class="sxs-lookup"><span data-stu-id="36862-333">4</span></span>|<span data-ttu-id="36862-334">有効</span><span class="sxs-lookup"><span data-stu-id="36862-334">Enabled</span></span>|<span data-ttu-id="36862-335">無効</span><span class="sxs-lookup"><span data-stu-id="36862-335">Disabled</span></span>|<span data-ttu-id="36862-336">新しいセッション</span><span class="sxs-lookup"><span data-stu-id="36862-336">New session</span></span>|<span data-ttu-id="36862-337">送信時アドインは実行されません。メッセージまたは会議アイテムは送信されます。</span><span class="sxs-lookup"><span data-stu-id="36862-337">No on-send add-ins run; message or meeting item is sent.</span></span>|<span data-ttu-id="36862-338">サポートされています。</span><span class="sxs-lookup"><span data-stu-id="36862-338">Supported.</span></span>|

#### <a name="web-browser-modern-outlook-windows-mac"></a><span data-ttu-id="36862-339">Web ブラウザー (モダン Outlook)、Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="36862-339">Web browser (modern Outlook), Windows, Mac</span></span>

<span data-ttu-id="36862-340">強制的に送信するには、管理者は両方のメールボックスでポリシーが有効になっていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="36862-340">To enforce on-send, administrators should ensure the policy has been enabled on both mailboxes.</span></span> <span data-ttu-id="36862-341">アドインで代理人アクセスをサポートする方法については、「共有フォルダーと共有メールボックスのシナリオを有効にする [」を参照してください](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="36862-341">To learn how to support delegate access in an add-in, see [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

### <a name="user-mailbox-with-on-send-add-in-featurepolicy-enabled-add-ins-that-support-on-send-are-installed-and-enabled-and-offline-mode-is-enabled"></a><span data-ttu-id="36862-342">ユーザー メールボックスで送信時アドイン機能/ポリシーが有効になっており、送信時機能をサポートするアドインがインストールされ、有効であり、オフライン モードが有効になっている</span><span class="sxs-lookup"><span data-stu-id="36862-342">User mailbox with on-send add-in feature/policy enabled, add-ins that support on-send are installed and enabled and offline mode is enabled</span></span>

<span data-ttu-id="36862-343">送信時アドインは、ユーザー、アドイン バックエンド、および Exchange のオンライン状態に従って実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-343">On-send add-ins will run according to the online state of the user, the add-in backend, and Exchange.</span></span>

#### <a name="users-state"></a><span data-ttu-id="36862-344">ユーザーの状態</span><span class="sxs-lookup"><span data-stu-id="36862-344">User's state</span></span>

<span data-ttu-id="36862-345">ユーザーがオンラインの場合、送信中に送信時アドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-345">The on-send add-ins will run during send if the user is online.</span></span> <span data-ttu-id="36862-346">ユーザーがオフラインの場合、送信中に送信時アドインは実行されず、メッセージまたは会議アイテムは送信されません。</span><span class="sxs-lookup"><span data-stu-id="36862-346">If the user is offline, the on-send add-ins will not run during send and the message or meeting item will not be sent.</span></span>

#### <a name="add-in-backends-state"></a><span data-ttu-id="36862-347">アドイン バックエンドの状態</span><span class="sxs-lookup"><span data-stu-id="36862-347">Add-in backend's state</span></span>

<span data-ttu-id="36862-348">送信時アドインは、バックエンドがオンラインで接続可能な場合に実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-348">An on-send add-in will run if its backend is online and reachable.</span></span> <span data-ttu-id="36862-349">バックエンドがオフラインの場合、送信は無効です。</span><span class="sxs-lookup"><span data-stu-id="36862-349">If the backend is offline, send is disabled.</span></span>

#### <a name="exchanges-state"></a><span data-ttu-id="36862-350">Exchange の状態</span><span class="sxs-lookup"><span data-stu-id="36862-350">Exchange's state</span></span>

<span data-ttu-id="36862-351">Exchange サーバーがオンラインでアクセスできる場合、送信中に送信時アドインが実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-351">The on-send add-ins will run during send if the Exchange server is online and reachable.</span></span> <span data-ttu-id="36862-352">送信時アドインが Exchange に接続できない場合、および該当するポリシーまたはコマンドレットが有効になっている場合、送信は無効です。</span><span class="sxs-lookup"><span data-stu-id="36862-352">If the on-send add-in cannot reach Exchange and the applicable policy or cmdlet is turned on, send is disabled.</span></span>

> [!NOTE]
> <span data-ttu-id="36862-353">オフライン状態の Mac では [**送信**] ボタン (または、既存の会議の場合は [**変更内容を送信**] ボタン) が無効になっており、ユーザーがオフラインの場合、組織が送信を許可していないという通知が表示されます。</span><span class="sxs-lookup"><span data-stu-id="36862-353">On Mac in any offline state, the **Send** button (or the **Send Update** button for existing meetings) is disabled and a notification displayed that their organization doesn't allow send when the user is offline.</span></span>

### <a name="user-can-edit-item-while-on-send-add-ins-are-working-on-it"></a><span data-ttu-id="36862-354">ユーザーは、オン送信アドインが作業している間にアイテムを編集できます</span><span class="sxs-lookup"><span data-stu-id="36862-354">User can edit item while on-send add-ins are working on it</span></span>

<span data-ttu-id="36862-355">送信時アドインがアイテムを処理している間、ユーザーは不適切なテキストや添付ファイルを追加してアイテムを編集できます。</span><span class="sxs-lookup"><span data-stu-id="36862-355">While on-send add-ins are processing an item, the user can edit the item by adding, for example, inappropriate text or attachments.</span></span> <span data-ttu-id="36862-356">アドインが送信時に処理されている間にユーザーがアイテムを編集するのを防ぐ場合は、ダイアログを使用して回避策を実装できます。</span><span class="sxs-lookup"><span data-stu-id="36862-356">If you want to prevent the user from editing the item while your add-in is processing on send, you can implement a workaround using a dialog.</span></span> <span data-ttu-id="36862-357">この回避策は、Outlook on the web (クラシック)、Windows Mac で使用できます。</span><span class="sxs-lookup"><span data-stu-id="36862-357">This workaround can be used in Outlook on the web (classic), Windows, and Mac.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="36862-358">モダン Outlook on the web: アドインの送信時の処理中にユーザーがアイテムを編集できない場合は、この記事の「送信時に使用する Outlook アドインのインストール」の説明に従って *OnSendAddinsEnabled* フラグを設定する必要があります。 `true` [](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send)</span><span class="sxs-lookup"><span data-stu-id="36862-358">Modern Outlook on the web: To prevent the user from editing the item while your add-in is processing on send, you should set the *OnSendAddinsEnabled* flag to `true` as described in the [Install Outlook add-ins that use on-send](outlook-on-send-addins.md?tabs=modern#install-outlook-add-ins-that-use-on-send) section earlier in this article.</span></span>

<span data-ttu-id="36862-359">送信時ハンドラーで、次の処理を行います。</span><span class="sxs-lookup"><span data-stu-id="36862-359">In your on-send handler:</span></span>

1. <span data-ttu-id="36862-360">[displayDialogAsync を](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-)呼び出してダイアログを開き、マウスのクリックとキーストロークが無効になります。</span><span class="sxs-lookup"><span data-stu-id="36862-360">Call [displayDialogAsync](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true#displaydialogasync-startaddress--options--callback-) to open a dialog so that mouse clicks and keystrokes are disabled.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="36862-361">クラシック モードでこの動作を[Outlook on the web、displayInIframe](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe)プロパティを呼び出しの `true` `options` パラメーターに設定する必要 `displayDialogAsync` があります。</span><span class="sxs-lookup"><span data-stu-id="36862-361">To get this behavior in classic Outlook on the web, you should set the [displayInIframe property](/javascript/api/office/office.dialogoptions?view=outlook-js-preview&preserve-view=true#displayiniframe) to `true` in the `options` parameter of the `displayDialogAsync` call.</span></span>

1. <span data-ttu-id="36862-362">アイテムの処理を実装します。</span><span class="sxs-lookup"><span data-stu-id="36862-362">Implement processing of the item.</span></span>
1. <span data-ttu-id="36862-363">ダイアログを閉じます。</span><span class="sxs-lookup"><span data-stu-id="36862-363">Close the dialog.</span></span> <span data-ttu-id="36862-364">また、ユーザーがダイアログを閉じるとどうなるかを処理します。</span><span class="sxs-lookup"><span data-stu-id="36862-364">Also, handle what happens if the user closes the dialog.</span></span>

## <a name="code-examples"></a><span data-ttu-id="36862-365">コード例</span><span class="sxs-lookup"><span data-stu-id="36862-365">Code examples</span></span>

<span data-ttu-id="36862-366">次のコード例は、単純な送信時アドインを作成する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="36862-366">The following code examples show you how to create a simple on-send add-in.</span></span> <span data-ttu-id="36862-367">これらの例を基にしたコード サンプルをダウンロードするには、「[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-367">To download the code sample that these examples are based on, see [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send).</span></span>

> [!TIP]
> <span data-ttu-id="36862-368">送信時イベントでダイアログを使用する場合は、イベントを完了する前にダイアログを閉じてください。</span><span class="sxs-lookup"><span data-stu-id="36862-368">If you use a dialog with the on-send event, make sure to close the dialog before completing the event.</span></span>

### <a name="manifest-version-override-and-event"></a><span data-ttu-id="36862-369">マニフェスト、バージョンのオーバーライド、イベント</span><span class="sxs-lookup"><span data-stu-id="36862-369">Manifest, version override, and event</span></span>

<span data-ttu-id="36862-370">[Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) コード サンプルには、2 つのマニフェストが含まれています。</span><span class="sxs-lookup"><span data-stu-id="36862-370">The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two manifests:</span></span>

- <span data-ttu-id="36862-371">`Contoso Message Body Checker.xml` &ndash; 制限された単語または機密情報についてメッセージの本文を送信時に確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="36862-371">`Contoso Message Body Checker.xml` &ndash; Shows how to check the body of a message for restricted words or sensitive information on send.</span></span>  

- <span data-ttu-id="36862-372">`Contoso Subject and CC Checker.xml` &ndash; CC 行に受信者を追加し、送信時にメッセージに件名が含まれていることを確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="36862-372">`Contoso Subject and CC Checker.xml` &ndash; Shows how to add a recipient to the CC line and verify that the message includes a subject line on send.</span></span>  

<span data-ttu-id="36862-373">`Contoso Message Body Checker.xml` マニフェスト ファイルには、`ItemSend` イベントで呼び出す必要がある関数ファイルと関数名を含めます。</span><span class="sxs-lookup"><span data-stu-id="36862-373">In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on the `ItemSend` event.</span></span> <span data-ttu-id="36862-374">操作は同期的に実行されます。</span><span class="sxs-lookup"><span data-stu-id="36862-374">The operation runs synchronously.</span></span>

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
> <span data-ttu-id="36862-375">Visual Studio 2019 を使用して送信時アドインを開発している場合は、次のような検証警告が表示される場合があります。"これは無効な xsi:type http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events ' 'です。この問題を回避するには、この警告に関するブログの GitHub gist として提供されている MailAppVersionOverridesV1_1.xsd の新しいバージョン[が必要です](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/)。</span><span class="sxs-lookup"><span data-stu-id="36862-375">If you are using Visual Studio 2019 to develop your on-send add-in, you may get a validation warning like the following: "This is an invalid xsi:type 'http://schemas.microsoft.com/office/mailappversionoverrides/1.1:Events'." To work around this, you'll need a newer version of the MailAppVersionOverridesV1_1.xsd which has been provided as a GitHub gist in a [blog about this warning](https://theofficecontext.com/2018/11/29/visual-studio-2017-this-is-an-invalid-xsitype-mailappversionoverrides-1-1event/).</span></span>

<span data-ttu-id="36862-376">`Contoso Subject and CC Checker.xml` マニフェスト ファイルの場合、次の例では、メッセージ送信イベントで呼び出す関数ファイルと関数名を示します。</span><span class="sxs-lookup"><span data-stu-id="36862-376">For the `Contoso Subject and CC Checker.xml` manifest file, the following example shows the function file and function name to call on message send event.</span></span>

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

<span data-ttu-id="36862-377">送信時 API には `VersionOverrides v1_1` が必要です。</span><span class="sxs-lookup"><span data-stu-id="36862-377">The on-send API requires `VersionOverrides v1_1`.</span></span> <span data-ttu-id="36862-378">マニフェストに `VersionOverrides` ノードを追加する方法を次に示します。</span><span class="sxs-lookup"><span data-stu-id="36862-378">The following shows you how to add the `VersionOverrides` node in your manifest.</span></span>

```xml
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
     <!-- On-send requires VersionOverridesV1_1 -->
     <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
         ...
     </VersionOverrides>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="36862-379">詳細については、以下を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-379">For more information, see the following:</span></span>
> - [<span data-ttu-id="36862-380">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="36862-380">Outlook add-in manifests</span></span>](manifests.md)
> - [<span data-ttu-id="36862-381">Office アドインの XML マニフェスト</span><span class="sxs-lookup"><span data-stu-id="36862-381">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)


### <a name="event-and-item-objects-and-bodygetasync-and-bodysetasync-methods"></a><span data-ttu-id="36862-382">`Event` オブジェクト、`item` オブジェクトと、`body.getAsync` メソッド、`body.setAsync` メソッドを理解する</span><span class="sxs-lookup"><span data-stu-id="36862-382">`Event` and `item` objects, and `body.getAsync` and `body.setAsync` methods</span></span>

<span data-ttu-id="36862-383">現在選択されているメッセージまたは会議アイテム (この例では、新しく作成されたメッセージ) にアクセスするには、`Office.context.mailbox.item` 名前空間を使用します。</span><span class="sxs-lookup"><span data-stu-id="36862-383">To access the currently selected message or meeting item (in this example, the newly composed message), use the `Office.context.mailbox.item` namespace.</span></span> <span data-ttu-id="36862-384">`ItemSend` イベントは、送信時機能によってマニフェストで指定された関数に自動的に渡されます &mdash; この例では `validateBody` 関数です。</span><span class="sxs-lookup"><span data-stu-id="36862-384">The `ItemSend` event is automatically passed by the on-send feature to the function specified in the manifest&mdash;in this example, the `validateBody` function.</span></span>

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

<span data-ttu-id="36862-385">`validateBody` 関数は、指定した形式 (HTML) の現在の本文を取得し、コールバック メソッドでのアクセスにコードが必要とする `ItemSend` イベント オブジェクトを渡します。</span><span class="sxs-lookup"><span data-stu-id="36862-385">The `validateBody` function gets the current body in the specified format (HTML) and passes the `ItemSend` event object that the code wants to access in the callback method.</span></span> <span data-ttu-id="36862-386">`getAsync` メソッドに加え、`Body` オブジェクトは本文を指定したテキストに置き換えるために使用できる `setAsync` メソッドも提供します。</span><span class="sxs-lookup"><span data-stu-id="36862-386">In addition to the `getAsync` method, the `Body` object also provides a `setAsync` method that you can use to replace the body with the specified text.</span></span>

> [!NOTE]
> <span data-ttu-id="36862-387">詳細については、「[Event オブジェクト](/javascript/api/office/office.addincommands.event)」と「[Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-387">For more information, see [Event Object](/javascript/api/office/office.addincommands.event) and [Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-).</span></span>
  

### <a name="notificationmessages-object-and-eventcompleted-method"></a><span data-ttu-id="36862-388">`NotificationMessages` オブジェクトと `event.completed` メソッド</span><span class="sxs-lookup"><span data-stu-id="36862-388">`NotificationMessages` object and `event.completed` method</span></span>

<span data-ttu-id="36862-389">`checkBodyOnlyOnSendCallBack` 関数は、正規表現を使用して、禁止された単語がメッセージの本文に含まれているかどうかを決定します。</span><span class="sxs-lookup"><span data-stu-id="36862-389">The `checkBodyOnlyOnSendCallBack` function uses a regular expression to determine whether the message body contains blocked words.</span></span> <span data-ttu-id="36862-390">制限されている単語の配列に対する一致が検出された場合、電子メールの送信をブロックし、情報バーを使用して送信者に通知します。</span><span class="sxs-lookup"><span data-stu-id="36862-390">If it finds a match against an array of restricted words, it then blocks the email from being sent and notifies the sender via the information bar.</span></span> <span data-ttu-id="36862-391">これを実行するには、`Item` オブジェクトの `notificationMessages` プロパティを使用して、`NotificationMessages` オブジェクトを返します。</span><span class="sxs-lookup"><span data-stu-id="36862-391">To do this, it uses the `notificationMessages` property of the `Item` object to return a `NotificationMessages` object.</span></span> <span data-ttu-id="36862-392">その後、次の例に示すように、`addAsync` メソッドを呼び出して通知をアイテムに追加します。</span><span class="sxs-lookup"><span data-stu-id="36862-392">It then adds a notification to the item by calling the `addAsync` method, as shown in the following example.</span></span>

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

<span data-ttu-id="36862-393">メソッドのパラメーターを次に示 `addAsync` します。</span><span class="sxs-lookup"><span data-stu-id="36862-393">The following are the parameters for the `addAsync` method.</span></span>

- <span data-ttu-id="36862-394">`NoSend` &ndash; 通知メッセージを参照するための開発者が指定したキーである文字列。</span><span class="sxs-lookup"><span data-stu-id="36862-394">`NoSend` &ndash; A string that is a developer-specified key to reference a notification message.</span></span> <span data-ttu-id="36862-395">これを使用して後でこのメッセージを変更できます。</span><span class="sxs-lookup"><span data-stu-id="36862-395">You can use it to modify this message later.</span></span> <span data-ttu-id="36862-396">キーは 32 文字を超えることはできません。</span><span class="sxs-lookup"><span data-stu-id="36862-396">The key can't be longer than 32 characters.</span></span>
- <span data-ttu-id="36862-397">`type` &ndash; JSON オブジェクト パラメーターのプロパティの 1 つ。</span><span class="sxs-lookup"><span data-stu-id="36862-397">`type` &ndash; One of the properties of the  JSON object parameter.</span></span> <span data-ttu-id="36862-398">メッセージの種類を表します。種類は [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) 列挙型の値に対応しています。</span><span class="sxs-lookup"><span data-stu-id="36862-398">Represents the type of a message; the types correspond to the values of the [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype) enumeration.</span></span> <span data-ttu-id="36862-399">使用可能な値は、進行状況のインジケーター、情報メッセージ、エラー メッセージです。</span><span class="sxs-lookup"><span data-stu-id="36862-399">Possible values are progress indicator, information message, or error message.</span></span> <span data-ttu-id="36862-400">この例では、`type` はエラー メッセージです。</span><span class="sxs-lookup"><span data-stu-id="36862-400">In this example, `type` is an error message.</span></span>  
- <span data-ttu-id="36862-401">`message` &ndash; JSON オブジェクト パラメーターのプロパティの 1 つ。</span><span class="sxs-lookup"><span data-stu-id="36862-401">`message` &ndash; One of the properties of the JSON object parameter.</span></span> <span data-ttu-id="36862-402">この例では、`message` は通知メッセージのテキストです。</span><span class="sxs-lookup"><span data-stu-id="36862-402">In this example, `message` is the text of the notification message.</span></span>

<span data-ttu-id="36862-403">アドインが送信操作によってトリガーされた `ItemSend` イベントの処理を完了したことを通知するには、`event.completed({allowEvent:Boolean})` メソッドを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="36862-403">To signal that the add-in has finished processing the `ItemSend` event triggered by the send operation, call the `event.completed({allowEvent:Boolean})` method.</span></span> <span data-ttu-id="36862-404">`allowEvent` プロパティは Boolean です。</span><span class="sxs-lookup"><span data-stu-id="36862-404">The `allowEvent` property is a Boolean.</span></span> <span data-ttu-id="36862-405">`true` に設定されている場合、送信が許可されます。</span><span class="sxs-lookup"><span data-stu-id="36862-405">If set to `true`, send is allowed.</span></span> <span data-ttu-id="36862-406">`false` に設定されている場合、電子メール メッセージの送信がブロックされます。</span><span class="sxs-lookup"><span data-stu-id="36862-406">If set to `false`, the email message is blocked from sending.</span></span>

> [!NOTE]
> <span data-ttu-id="36862-407">詳細については、「[notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)」と「[completed](/javascript/api/office/office.addincommands.event)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-407">For more information, see [notificationMessages](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [completed](/javascript/api/office/office.addincommands.event).</span></span>

### <a name="replaceasync-removeasync-and-getallasync-methods"></a><span data-ttu-id="36862-408">`replaceAsync` メソッド、`removeAsync` メソッド、および`getAllAsync`メソッド</span><span class="sxs-lookup"><span data-stu-id="36862-408">`replaceAsync`, `removeAsync`, and `getAllAsync` methods</span></span>

<span data-ttu-id="36862-409">`addAsync` メソッドに加え、`NotificationMessages` オブジェクトは本文を指定したテキストに置き換えるために使用できる `replaceAsync`、`removeAsync`、および `getAllAsync` の各メソッドも提供します。</span><span class="sxs-lookup"><span data-stu-id="36862-409">In addition to the `addAsync` method, the `NotificationMessages` object also includes `replaceAsync`, `removeAsync`, and `getAllAsync` methods.</span></span>  <span data-ttu-id="36862-410">このコード サンプルでは、これらのメソッドは使用されません。</span><span class="sxs-lookup"><span data-stu-id="36862-410">These methods are not used in this code sample.</span></span>  <span data-ttu-id="36862-411">詳細については、「[NotificationMessages](/javascript/api/outlook/office.NotificationMessages)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="36862-411">For more information, see [NotificationMessages](/javascript/api/outlook/office.NotificationMessages).</span></span>


### <a name="subject-and-cc-checker-code"></a><span data-ttu-id="36862-412">件名および CC のチェッカー コード</span><span class="sxs-lookup"><span data-stu-id="36862-412">Subject and CC checker code</span></span>

<span data-ttu-id="36862-413">次のコード例では、CC 行に受信者を追加し、送信時にメッセージに件名が含まれていることを確認する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="36862-413">The following code example shows you how to add a recipient to the CC line and verify that the message includes a subject on send.</span></span> <span data-ttu-id="36862-414">この例では、送信時機能を使用して、電子メールの送信を許可または禁止します。</span><span class="sxs-lookup"><span data-stu-id="36862-414">This example uses the on-send feature to allow or disallow an email from sending.</span></span>  

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

<span data-ttu-id="36862-p155">CC 行に受信者を追加して、送信時にメッセージに件名が含まれていることを確認する方法、および使用可能な API を表示する方法の詳細については、「[Outlook-Add-in-On-Send サンプル](https://github.com/OfficeDev/Outlook-Add-in-On-Send)」を参照してください。コードには詳細なコメントが付けられています。</span><span class="sxs-lookup"><span data-stu-id="36862-p155">To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.</span></span>

## <a name="see-also"></a><span data-ttu-id="36862-417">関連項目</span><span class="sxs-lookup"><span data-stu-id="36862-417">See also</span></span>

- [<span data-ttu-id="36862-418">Outlook アドインのアーキテクチャと機能の概要</span><span class="sxs-lookup"><span data-stu-id="36862-418">Overview of Outlook add-ins architecture and features</span></span>](outlook-add-ins-overview.md)
- [<span data-ttu-id="36862-419">アドイン コマンド デモの Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="36862-419">Add-in Command Demo Outlook add-in</span></span>](https://github.com/OfficeDev/outlook-add-in-command-demo)