---
title: Outlook コンテキスト アドインのアクティブ化のトラブルシューティング
description: アドインが期待どおりにアクティブにならない場合は、考えられる理由について、次の点を調査してください。
ms.date: 05/27/2020
localization_priority: Normal
ms.openlocfilehash: 555ae2a45bf49d74d1fd439258fd87035644e86a
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413183"
---
# <a name="troubleshoot-outlook-add-in-activation"></a><span data-ttu-id="16941-103">Outlook アドインのアクティブ化のトラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="16941-103">Troubleshoot Outlook add-in activation</span></span>

<span data-ttu-id="16941-p101">Outlook コンテキスト アドインのアクティブ化は、アドイン マニフェスト内のアクティブ化ルールに基づいて行われます。現在選択されているアイテムの条件がアドインのアクティブ化ルールを満たす場合に、ホスト アプリケーションは、Outlook の UI (新規作成アドインではアドイン選択ウィンドウ、閲覧アドインではアドイン バー) で [アドイン] ボタンをアクティブ化して表示します。ただし、アドインが期待どおりにアクティブにならない場合は、考えられる理由について、次の点を調査してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p101">Outlook contextual add-in activation is based on the activation rules in the add-in manifest. When conditions for the currently selected item satisfy the activation rules for the add-in, the host application activates and displays the add-in button in the Outlook UI (add-in selection pane for compose add-ins, add-in bar for read add-ins). However, if your add-in doesn't activate as you expect, you should look into the following areas for possible reasons.</span></span>

## <a name="is-user-mailbox-on-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a><span data-ttu-id="16941-107">ユーザーのメールボックスが、Exchange 2013 以降のバージョンの Exchange Server 上にあるか?</span><span class="sxs-lookup"><span data-stu-id="16941-107">Is user mailbox on a version of Exchange Server that is at least Exchange 2013?</span></span>

<span data-ttu-id="16941-p102">まず、テストしているユーザーの電子メール アカウントが、Exchange 2013 以降のバージョンの Exchange Server 上にあることを確認します。Exchange 2013 より後にリリースされた特定の機能を使用する場合は、ユーザーのアカウントが Exchange の適切なバージョン上にあることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p102">First, ensure that the user's email account you're testing with is on a version of Exchange Server that is at least Exchange 2013. If you are using specific features that are released after Exchange 2013, make sure the user's account is on the appropriate version of Exchange.</span></span>

<span data-ttu-id="16941-110">Exchange 2013 のバージョンは、次の方法のいずれかを使用して確認できます。</span><span class="sxs-lookup"><span data-stu-id="16941-110">You can verify the version of Exchange 2013 by using one of the following approaches:</span></span>

- <span data-ttu-id="16941-111">Exchange Server 管理者に確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-111">Check with your Exchange Server administrator.</span></span>

- <span data-ttu-id="16941-p103">スクリプト デバッガー (たとえば、Internet Explorer に付属する JScript デバッガーなど) で Outlook on the web またはモバイル デバイス上のアドインをテストしている場合は、スクリプトの読み込み元を指定する **script** タグの **src** 属性を探します。このパスには、**owa/15.0.516.x/owa2/...** という部分文字列があります。この中の **15.0.516.x** が Exchange Server のバージョン (**15.0.516.2** など) を表します。</span><span class="sxs-lookup"><span data-stu-id="16941-p103">If you are testing the add-in on Outlook on the web or mobile devices, in a script debugger (for example, the JScript Debugger that comes with Internet Explorer), look for the **src** attribute of the **script** tag that specifies the location from which scripts are loaded. The path should contain a substring **owa/15.0.516.x/owa2/...**, where **15.0.516.x** represents the version of the Exchange Server, such as **15.0.516.2**.</span></span>

- <span data-ttu-id="16941-p104">あるいは、[Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) プロパティを使用してバージョンを確認することもできます。Outlook on the web およびモバイル デバイス上で、このプロパティは Exchange Server のバージョンを返します。</span><span class="sxs-lookup"><span data-stu-id="16941-p104">Alternatively, you can use the [Office.context.mailbox.diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion) property to verify the version. On Outlook on the web and mobile devices, this property returns the version of the Exchange Server.</span></span>

- <span data-ttu-id="16941-116">Outlook 上でアドインをテストできる場合は、次に示す Outlook オブジェクト モデルと Visual Basic エディターを使用した簡単なデバッグ方法を使用できます。</span><span class="sxs-lookup"><span data-stu-id="16941-116">If you can test the add-in on Outlook, you can use the following simple debugging technique that uses the Outlook object model and Visual Basic Editor:</span></span>

    1. <span data-ttu-id="16941-p105">最初に、Outlook でマクロが有効になっていることを確認します。**[ファイル]**、**[オプション]**、**[セキュリティ センター]**、**[セキュリティ センターの設定]**、**[マクロの設定]** の順に選択します。セキュリティ センターで、**[すべてのマクロの通知]** が選択されていることを確認します。Outlook の起動時に **[マクロを有効にする]** も選択している必要があります。</span><span class="sxs-lookup"><span data-stu-id="16941-p105">First, verify that macros are enabled for Outlook. Choose **File**, **Options**, **Trust Center**, **Trust Center Settings**, **Macro Settings**. Ensure that **Notifications for all macros** is selected in the Trust Center. You should have also selected **Enable Macros** during Outlook startup.</span></span>

    1. <span data-ttu-id="16941-121">リボンの **[開発]** タブで **[Visual Basic]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="16941-121">On the **Developer** tab of the ribbon, choose **Visual Basic**.</span></span>

       > [!NOTE]
       > <span data-ttu-id="16941-p106">**[開発]** タブが表示されない場合には、「[方法:[開発] タブをリボンに表示する](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon)」を参照して、有効にしてください。</span><span class="sxs-lookup"><span data-stu-id="16941-p106">Not seeing the **Developer** tab? See [How to: Show the Developer Tab on the Ribbon](/visualstudio/vsto/how-to-show-the-developer-tab-on-the-ribbon) to turn it on.</span></span>

    1. <span data-ttu-id="16941-124">Visual Basic エディターで、**[表示]**、**[イミディエイト ウィンドウ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="16941-124">In the Visual Basic Editor, choose **View**, **Immediate Window**.</span></span>

    1. <span data-ttu-id="16941-p107">イミディエイト ウィンドウに次のように入力し、Exchange Server のバージョンを表示します。戻される値のメジャー バージョンは、15 以上である必要があります。</span><span class="sxs-lookup"><span data-stu-id="16941-p107">Type the following in the Immediate window to display the version of the Exchange Server. The major version of the returned value must be equal to or greater than 15.</span></span>

       - <span data-ttu-id="16941-127">ユーザーのプロファイルに Exchange アカウントが 1 つだけある場合:</span><span class="sxs-lookup"><span data-stu-id="16941-127">If there is only one Exchange account in the user's profile:</span></span>

       ```vb
        ?Session.ExchangeMailboxServerVersion
       ```

       - <span data-ttu-id="16941-128">同じユーザー プロファイルに複数の Exchange アカウントがある場合 (`emailAddress` はユーザーのプライマリ STMP アドレスを含む文字列を表します):</span><span class="sxs-lookup"><span data-stu-id="16941-128">If there are multiple Exchange accounts in the same user profile (`emailAddress` represents a string that contains the user's primary SMTP address):</span></span>

       ```vb
        ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
       ```

## <a name="is-the-add-in-disabled"></a><span data-ttu-id="16941-129">アドインが無効化されていないか?</span><span class="sxs-lookup"><span data-stu-id="16941-129">Is the add-in disabled?</span></span>

<span data-ttu-id="16941-p108">いずれかの Outlook リッチ クライアントで、パフォーマンス上の理由によりアドインが無効化されている可能性があります。たとえば、CPU コア使用率やメモリ使用量のしきい値、クラッシュ許容度、およびアドインに対するすべての正規表現の処理時間が超過した場合などです。このようなことが起きると、Outlook リッチ クライアントは、アドインを無効化していることを示す通知を表示します。</span><span class="sxs-lookup"><span data-stu-id="16941-p108">Any one of the Outlook rich clients can disable an add-in for performance reasons, including exceeding usage thresholds for CPU core or memory, tolerance for crashes, and length of time to process all the regular expressions for an add-in. When this happens, the Outlook rich client displays a notification that it is disabling the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="16941-132">リソース使用量を監視するのは Outlook リッチ クライアントだけですが、Outlook リッチ クライアントでアドインを無効化すると、Outlook on the web とモバイル デバイスでもアドインが無効化されます。</span><span class="sxs-lookup"><span data-stu-id="16941-132">Only Outlook rich clients monitor resource usage, but disabling an add-in in an Outlook rich client also disables the add-in in Outlook on the web and mobile devices.</span></span>

<span data-ttu-id="16941-133">次のどちらかの方法を使用して、アドインが無効化されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-133">Use one of the following approaches to verify whether an add-in is disabled:</span></span>

- <span data-ttu-id="16941-134">Outlook on the web の場合、電子メール アカウントに直接サインインして、[設定] アイコンを選択し、**[アドインの管理]** を選択して、Exchange 管理センターにアクセスします。ここで、アドインが有効化されているかどうかを確認できます。</span><span class="sxs-lookup"><span data-stu-id="16941-134">In Outlook on the web, sign in directly to the email account, choose the Settings icon, and then choose **Manage add-ins** to go to the Exchange Admin Center, where you can verify whether the add-in is enabled.</span></span>

- <span data-ttu-id="16941-135">Windows 用 Outlook の場合、Backstage ビューに移動し、**[アドインの管理]** を選択します。それから、Exchange 管理センターにサインインし、アドインが有効化されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-135">In Outlook on Windows, go to the Backstage view and choose **Manage add-ins**. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.</span></span>

- <span data-ttu-id="16941-p109">Mac 用 Outlook の場合は、アドイン バーで **[アドインの管理]** を選択します。それから、Exchange 管理センターにサインインし、アドインが有効化されているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-p109">In Outlook on Mac, choose **Manage add-ins** in the add-in bar. Sign in to the Exchange Admin Center to verify whether the add-in is enabled.</span></span>

## <a name="does-the-tested-item-support-outlook-add-ins-is-the-selected-item-delivered-by-a-version-of-exchange-server-that-is-at-least-exchange-2013"></a><span data-ttu-id="16941-p110">テストするアイテムが Outlook アドインをサポートしているか? 選択されたアイテムが Exchange 2013 以降のバージョンの Exchange Server で配信されているか?</span><span class="sxs-lookup"><span data-stu-id="16941-p110">Does the tested item support Outlook add-ins? Is the selected item delivered by a version of Exchange Server that is at least Exchange 2013?</span></span>

<span data-ttu-id="16941-140">Outlook アドインが閲覧アドインであり、ユーザーがメッセージ (メール メッセージ、会議出席依頼、返信、キャンセルなど) や予定を表示するときにアクティブ化されるものである場合、これらのアイテムが通常はアドインをサポートしているとしても、選択しているアイテムが次のいずれかの場合は例外があります。</span><span class="sxs-lookup"><span data-stu-id="16941-140">If your Outlook add-in is a read add-in and is supposed to be activated when the user is viewing a message (including email messages, meeting requests, responses, and cancellations) or appointment, even though these items generally support add-ins, there are exceptions.</span></span> <span data-ttu-id="16941-141">選択したアイテムが[アクティブではない Outlook アドインの一覧](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)にあるかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-141">Check if the selected item is one of those [listed where Outlook add-ins do not activate](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins).</span></span>

<span data-ttu-id="16941-142">また、予定は常にリッチ テキスト形式で保存されるので、[BodyAsHTML](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) の **PropertyName** 値を指定する **ItemHasRegularExpressionMatch** ルールでは、プレーン テキストやリッチ テキスト形式で保存された予定またはメッセージ上でアドインがアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="16941-142">Also, because appointments are always saved in Rich Text Format, an [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rule that specifies a **PropertyName** value of **BodyAsHTML** would not activate an add-in on an appointment or message that is saved in plain text or Rich Text Format.</span></span>

<span data-ttu-id="16941-p112">メール アイテムが上記の種類のいずれかでなくても、アイテムが Exchange 2013 以降のバージョンの Exchange Server で配信されたものでない場合、そのアイテムでは、送信者の SMTP アドレスなどの既知のエンティティおよびプロパティが識別できません。これらのエンティティやプロパティに依存するアクティブ化ルールはどれも条件が満たされず、そのアドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="16941-p112">Even if a mail item is not one of the above types, if the item was not delivered by a version of Exchange Server that is at least Exchange 2013, known entities and properties such as sender's SMTP address would not be identified on the item. Any activation rules that rely on these entities or properties would not be satisfied, and the add-in would not be activated.</span></span>

<span data-ttu-id="16941-145">アドインが新規作成アドインであり、ユーザーがメッセージや会議出席依頼を作成するときにアクティブ化されるものである場合、そのアイテムが IRM によって保護されていないことを確認してください。</span><span class="sxs-lookup"><span data-stu-id="16941-145">If your add-in is a compose add-in and is supposed to be activated when the user is authoring a message or meeting request, make sure the item is not protected by IRM.</span></span>

## <a name="is-the-add-in-manifest-installed-properly-and-does-outlook-have-a-cached-copy"></a><span data-ttu-id="16941-146">アドイン マニフェストが適切にインストールされているか? また Outlook にキャッシュ コピーがあるか?</span><span class="sxs-lookup"><span data-stu-id="16941-146">Is the add-in manifest installed properly, and does Outlook have a cached copy?</span></span>

<span data-ttu-id="16941-p113">このシナリオは Windows での Outlook にのみ適用されます。通常、メールボックスに Outlook アドインをインストールすると、Exchange Server は、アドイン マニフェストを指定の場所からその Exchange Server 上のメールボックスにコピーします。Outlook は起動するたびに、そのメールボックスにインストールされたすべてのマニフェストを、次の場所にある一時的なキャッシュに読み込みます。</span><span class="sxs-lookup"><span data-stu-id="16941-p113">This scenario applies to only Outlook on Windows. Normally, when you install an Outlook add-in for a mailbox, the Exchange Server copies the add-in manifest from the location you indicate to the mailbox on that Exchange Server. Every time Outlook starts, it reads all the manifests installed for that mailbox into a temporary cache at the following location:</span></span>

```text
%LocalAppData%\Microsoft\Office\16.0\WEF
```

<span data-ttu-id="16941-150">たとえば、ユーザー John の場合、キャッシュは C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF. にある可能性があります。</span><span class="sxs-lookup"><span data-stu-id="16941-150">For example, for the user John, the cache might be at C:\Users\john\AppData\Local\Microsoft\Office\16.0\WEF.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="16941-151">Windows の Outlook 2013 では、16.0 ではなく15.0 を使用して、場所を次のようにします。</span><span class="sxs-lookup"><span data-stu-id="16941-151">For Outlook 2013 on Windows, use 15.0 instead of 16.0 so the location would be:</span></span>
>
> ```text
> %LocalAppData%\Microsoft\Office\15.0\WEF
> ```

<span data-ttu-id="16941-p114">アドインがどのアイテムに対してもアクティブ化されない場合、マニフェストが Exchange Server 上に適切にインストールされなかったか、あるいは、Outlook が起動時に正しくマニフェストを読み取れなかった可能性があります。Exchange 管理センターを使用して、アドインがメールボックスにインストールされ、有効化されていることを確認し、必要に応じて Exchange Server を再起動します。</span><span class="sxs-lookup"><span data-stu-id="16941-p114">If an add-in does not activate for any items, the manifest might not have been installed properly on the Exchange Server, or Outlook has not read the manifest properly on startup. Using the Exchange Admin Center, ensure that the add-in is installed and enabled for your mailbox, and reboot the Exchange Server, if necessary.</span></span>

<span data-ttu-id="16941-154">図 1 は、Outlook に有効なバージョンのマニフェストがあるかどうかを確認するステップの概要を示しています。</span><span class="sxs-lookup"><span data-stu-id="16941-154">Figure 1 shows a summary of the steps to verify whether Outlook has a valid version of the manifest.</span></span>

<span data-ttu-id="16941-155">**図 1.Outlook がマニフェストを適切にキャッシュしたかどうかを確認するステップのフローチャート**</span><span class="sxs-lookup"><span data-stu-id="16941-155">**Figure 1. Flow chart of the steps to verify whether Outlook properly cached the manifest**</span></span>

![マニフェストを確認するためのフローチャート](../images/troubleshoot-manifest-flow.png)

<span data-ttu-id="16941-157">以下の手順では、その詳細を説明します。</span><span class="sxs-lookup"><span data-stu-id="16941-157">The following procedure describes the details.</span></span>

1. <span data-ttu-id="16941-158">Outlook を開いている間にマニフェストを変更し、アドインの開発に Visual Studio 2012 や Visual Studio の新しいバージョンを使用していない場合は、Exchange 管理センターを使用して、そのアドインをアンインストールし、再インストールする必要があります。</span><span class="sxs-lookup"><span data-stu-id="16941-158">If you have modified the manifest while Outlook is open, and you are not using Visual Studio 2012 or a later version of Visual Studio to develop the add-in, you should uninstall the add-in and reinstall it using the Exchange Admin Center.</span></span>

1. <span data-ttu-id="16941-159">Outlook を再起動し、Outlook でアドインがアクティブになっているかどうかをテストします。</span><span class="sxs-lookup"><span data-stu-id="16941-159">Restart Outlook and test whether Outlook now activates the add-in.</span></span>

1. <span data-ttu-id="16941-p115">アドインがアクティブ化されない場合は、アドインのマニフェストの適切なキャッシュ コピーが Outlook にあるかどうかを確認します。次のパスの下を探してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p115">If Outlook doesn't activate the add-in, check whether Outlook has a properly cached copy of the manifest for the add-in. Look under the following path:</span></span>

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF
    ```

    <span data-ttu-id="16941-162">次のサブフォルダーでマニフェストを見つけることができます。</span><span class="sxs-lookup"><span data-stu-id="16941-162">You can find the manifest in the following subfolder:</span></span>

    ```text
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
    ```

    > [!NOTE]
    > <span data-ttu-id="16941-163">ユーザー John のメールボックスにインストールされたマニフェストへのパスの例は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="16941-163">The following is an example of a path to a manifest installed for a mailbox for the user John:</span></span>
    >
    > ```text
    > C:\Users\john\appdata\Local\Microsoft\Office\16.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    > ```

    <span data-ttu-id="16941-164">テストしているアドインのマニフェストが、キャッシュされたマニフェストに含まれているかどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-164">Verify whether the manifest of the add-in you're testing is among the cached manifests.</span></span>

1. <span data-ttu-id="16941-165">マニフェストがキャッシュにある場合は、このセクションの残りの部分をスキップして、このセクションの後で説明している、他に考えられる理由を検討します。</span><span class="sxs-lookup"><span data-stu-id="16941-165">If the manifest is in the cache, skip the rest of this section and consider the other possible reasons following this section.</span></span>

1. <span data-ttu-id="16941-p116">マニフェストがキャッシュにない場合は、Outlook が Exchange Server から実際にマニフェストを読み取ったかどうかを確認します。これを行うには、Windows イベント ビューアーを使用します。</span><span class="sxs-lookup"><span data-stu-id="16941-p116">If the manifest is not in the cache, check whether Outlook indeed successfully read the manifest from the Exchange Server. To do that, use the Windows Event Viewer:</span></span>

    1. <span data-ttu-id="16941-168">**[Windows ログ]** で **[アプリケーション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="16941-168">Under **Windows Logs**, choose **Application**.</span></span>

    1. <span data-ttu-id="16941-169">イベント ID が 63 に等しい比較的最近のイベントを探します。これは、Outlook が Exchange Server からマニフェストをダウンロードしたことを表します。</span><span class="sxs-lookup"><span data-stu-id="16941-169">Look for a reasonably recent event for which the Event ID equals 63, which represents Outlook downloading a manifest from an Exchange Server.</span></span>

    1. <span data-ttu-id="16941-170">Outlook によるマニフェストの読み取りが正常に行われた場合は、記録されたイベントに次の説明があります。</span><span class="sxs-lookup"><span data-stu-id="16941-170">If Outlook successfully read a manifest, the logged event should have the following description:</span></span>

        ```text
        The Exchange web service request GetAppManifests succeeded.
        ```

        <span data-ttu-id="16941-171">このセクションの残りの部分をスキップして、このセクションの後で説明している、他に考えられる理由を検討します。</span><span class="sxs-lookup"><span data-stu-id="16941-171">Then skip the rest of this section and consider the other possible reasons following this section.</span></span>

1. <span data-ttu-id="16941-172">イベントの成功を確認できない場合は、Outlook を閉じて、次のパスにあるすべてのマニフェストを削除します。</span><span class="sxs-lookup"><span data-stu-id="16941-172">If you don't see a successful event, close Outlook, and delete all the manifests in the following path:</span></span>

    ```text
    %LocalAppData%\Microsoft\Office\16.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
    ```

    <span data-ttu-id="16941-173">Outlook を起動し、Outlook でアドインがアクティブになっているかどうかをテストします。</span><span class="sxs-lookup"><span data-stu-id="16941-173">Start Outlook and test whether Outlook now activates the add-in.</span></span>

1. <span data-ttu-id="16941-174">アドインがアクティブ化されない場合は、手順 3 に戻り、Outlook がマニフェストを適切に読み取ったかどうかを再度確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-174">If Outlook doesn't activate the add-in, go back to Step 3 to verify again whether Outlook has properly read the manifest.</span></span>

## <a name="is-the-add-in-manifest-valid"></a><span data-ttu-id="16941-175">アドイン マニフェストは有効か?</span><span class="sxs-lookup"><span data-stu-id="16941-175">Is the add-in manifest valid?</span></span>

<span data-ttu-id="16941-176">「[マニフェストの問題を検証し、トラブルシューティングを行う](../testing/troubleshoot-manifest.md)」を参照して、アドインのマニフェストの問題をデバッグしてください。</span><span class="sxs-lookup"><span data-stu-id="16941-176">See [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>

## <a name="are-you-using-the-appropriate-activation-rules"></a><span data-ttu-id="16941-177">適切なアクティブ化ルールを使用しているか?</span><span class="sxs-lookup"><span data-stu-id="16941-177">Are you using the appropriate activation rules?</span></span>

<span data-ttu-id="16941-p117">Office アドイン マニフェスト スキーマ バージョン 1.1 以降では、ユーザーが新規作成フォームを使用しているときにアクティブ化されるアドイン (新規作成アドイン) や閲覧フォームを使用しているときにアクティブ化されるアドイン (閲覧アドイン) を作成できます。アドインをアクティブ化するフォームの種類に適した正しいアクティブ化ルールを指定してください。たとえば、新規作成アドインをアクティブ化する場合は、[FormType](../reference/manifest/rule.md#itemis-rule) 属性が **Edit** または **ReadOrEdit** に設定された **ItemIs** ルールのみを使用する必要があり、[ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールや [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ルールなど他の型のルールを新規作成アドイン用に使用することはできません。詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p117">Starting in version 1.1 of the Office Add-ins manifests schema, you can create add-ins that are activated when the user is in a compose form (compose add-ins) or in a read form (read add-ins). Make sure you specify the appropriate activation rules for each type of form that your add-in is supposed to activate in. For example, you can activate compose add-ins using only [ItemIs](../reference/manifest/rule.md#itemis-rule) rules with the **FormType** attribute set to **Edit** or **ReadOrEdit**, and you cannot use any of the other types of rules, such as [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) and [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) rules for compose add-ins. For more information, see [Activation rules for Outlook add-ins](activation-rules.md).</span></span>

## <a name="if-you-use-a-regular-expression-is-it-properly-specified"></a><span data-ttu-id="16941-181">正規表現を使用している場合、正しく指定されていますか。</span><span class="sxs-lookup"><span data-stu-id="16941-181">If you use a regular expression, is it properly specified?</span></span>

<span data-ttu-id="16941-p118">アクティブ化ルール内の正規表現は閲覧アドインの XML マニフェスト ファイルの一部であるため、正規表現で特定の文字を使用する場合は、XML プロセッサがサポートする対応するエスケープ シーケンスに従う必要があります。表 1 にこのような特殊文字を示します。</span><span class="sxs-lookup"><span data-stu-id="16941-p118">Because regular expressions in activation rules are part of the XML manifest file for a read add-in, if a regular expression uses certain characters, be sure to follow the corresponding escape sequence that XML processors support. Table 1 lists these special characters.</span></span>

<span data-ttu-id="16941-184">**表 1.正規表現のエスケープ シーケンス**</span><span class="sxs-lookup"><span data-stu-id="16941-184">**Table 1. Escape sequences for regular expressions**</span></span>

|<span data-ttu-id="16941-185">**文字**</span><span class="sxs-lookup"><span data-stu-id="16941-185">**Character**</span></span>|<span data-ttu-id="16941-186">**説明**</span><span class="sxs-lookup"><span data-stu-id="16941-186">**Description**</span></span>|<span data-ttu-id="16941-187">**使用するエスケープ シーケンス**</span><span class="sxs-lookup"><span data-stu-id="16941-187">**Escape sequence to use**</span></span>|
|:-----|:-----|:-----|
|`"`|<span data-ttu-id="16941-188">二重引用符</span><span class="sxs-lookup"><span data-stu-id="16941-188">Double quotation mark</span></span>|<span data-ttu-id="16941-189">&amp;quot;</span><span class="sxs-lookup"><span data-stu-id="16941-189">&amp;quot;</span></span>|
|`&`|<span data-ttu-id="16941-190">アンパサンド</span><span class="sxs-lookup"><span data-stu-id="16941-190">Ampersand</span></span>|<span data-ttu-id="16941-191">&amp;amp;</span><span class="sxs-lookup"><span data-stu-id="16941-191">&amp;amp;</span></span>|
|`'`|<span data-ttu-id="16941-192">アポストロフィ</span><span class="sxs-lookup"><span data-stu-id="16941-192">Apostrophe</span></span>|<span data-ttu-id="16941-193">&amp;apos;</span><span class="sxs-lookup"><span data-stu-id="16941-193">&amp;apos;</span></span>|
|`<`|<span data-ttu-id="16941-194">より小さい</span><span class="sxs-lookup"><span data-stu-id="16941-194">Less-than sign</span></span>|<span data-ttu-id="16941-195">&amp;lt;</span><span class="sxs-lookup"><span data-stu-id="16941-195">&amp;lt;</span></span>|
|`>`|<span data-ttu-id="16941-196">より大きい</span><span class="sxs-lookup"><span data-stu-id="16941-196">Greater-than sign</span></span>|<span data-ttu-id="16941-197">&amp;gt;</span><span class="sxs-lookup"><span data-stu-id="16941-197">&amp;gt;</span></span>|

## <a name="if-you-use-a-regular-expression-is-the-read-add-in-activating-in-outlook-on-the-web-or-mobile-devices-but-not-in-any-of-the-outlook-rich-clients"></a><span data-ttu-id="16941-198">正規表現を使用する場合、閲覧アドインは Outlook on the web またはモバイル デバイスではアクティブ化されるものの、どの Outlook リッチ クライアントでもアクティブ化されないか?</span><span class="sxs-lookup"><span data-stu-id="16941-198">If you use a regular expression, is the read add-in activating in Outlook on the web or mobile devices, but not in any of the Outlook rich clients?</span></span>

<span data-ttu-id="16941-p119">Outlook リッチ クライアントでは、Outlook on the web とモバイル デバイスで使用されている正規表現エンジンとでは、異なる正規表現エンジンを使用します。Outlook リッチ クライアントでは、Visual Studio の標準テンプレート ライブラリの一部として提供されている C++ 正規表現エンジンを使用します。このエンジンは ECMAScript 5 標準に準拠しています。Outlook on the web およびモバイル デバイスでは、JavaScript の一部である正規表現評価を使用します。これはブラウザーによって提供されるものであり、ECMAScript 5 のスーパーセットをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="16941-p119">Outlook rich clients use a regular expression engine that's different from the one used by Outlook on the web and mobile devices. Outlook rich clients use the C++ regular expression engine provided as part of the Visual Studio standard template library. This engine complies with ECMAScript 5 standards. Outlook on the web and mobile devices use regular expression evaluation that is part of JavaScript, is provided by the browser, and supports a superset of ECMAScript 5.</span></span>

<span data-ttu-id="16941-p120">ほとんどの場合、これらのホスト アプリケーションはアクティブ化ルールでの同じ正規表現に対して同じ一致を検出しますが、例外もあります。たとえば、正規表現に事前定義された文字クラスに基づいたカスタム文字のクラスが含まれている場合、Outlook リッチ クライアントは Outlook on the web やモバイル デバイスとは異なる結果を返す可能性があります。一例を挙げると、文字クラス内に短縮形の文字クラス `[\d\w]` が含まれる場合は、異なる結果を返します。この場合、異なるホストでの異なる結果を回避するには `(\d|\w)` を使用します。</span><span class="sxs-lookup"><span data-stu-id="16941-p120">While in most cases, these host applications find the same matches for the same regular expression in an activation rule, there are exceptions. For instance, if the regex includes a custom character class based on predefined character classes, an Outlook rich client may return results different from Outlook on the web and mobile devices. As an example, character classes that contain shorthand character classes  `[\d\w]` within them would return different results. In this case, to avoid different results on different hosts, use `(\d|\w)` instead.</span></span>

<span data-ttu-id="16941-p121">正規表現を十分にテストしてください。異なる結果が返された場合は、両方のエンジンでの互換性のために正規表現を書き換えます。Outlook リッチ クライアントの評価結果を確認するには、一致させるテキストのサンプルに対して正規表現を適用させる小さな C++ プログラムを作成します。Visual Studio 上で動作する C++ テスト プログラムは、標準テンプレート ライブラリを使用して、同じ正規表現を実行しているときに Outlook リッチ クライアントの動作をシミュレートします。Outlook on the web またはモバイル デバイスでの評価結果を確認するには、お好きな JavaScript 正規表現テスターを使用してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p121">Test your regular expression thoroughly. If it returns different results, rewrite the regular expression for compatibility with both engines. To verify evaluation results on an Outlook rich client, write a small C++ program that applies the regular expression against a sample of the text you are trying to match. Running on Visual Studio, the C++ test program would use the standard template library, simulating the behavior of the Outlook rich client when running the same regular expression. To verify evaluation results on Outlook on the web or mobile devices, use your favorite JavaScript regular expression tester.</span></span>

## <a name="if-you-use-an-itemis-itemhasattachment-or-itemhasregularexpressionmatch-rule-have-you-verified-the-related-item-property"></a><span data-ttu-id="16941-212">ItemIs ルール、ItemHasAttachment ルール、または ItemHasRegularExpressionMatch ルールを使用する場合、関連するアイテム プロパティを確認しましたか。</span><span class="sxs-lookup"><span data-stu-id="16941-212">If you use an ItemIs, ItemHasAttachment, or ItemHasRegularExpressionMatch rule, have you verified the related item property?</span></span>

<span data-ttu-id="16941-213">**ItemHasRegularExpressionMatch** アクティブ化ルールを使用する場合は、**PropertyName** 属性の値が、選択されているアイテムの予期する値かどうかを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-213">If you use an **ItemHasRegularExpressionMatch** activation rule, verify whether the value of the **PropertyName** attribute is what you expect for the selected item.</span></span> <span data-ttu-id="16941-214">対応するプロパティをデバッグするときのいくつかのヒントを次に示します。</span><span class="sxs-lookup"><span data-stu-id="16941-214">The following are some tips to debug the corresponding properties:</span></span>

- <span data-ttu-id="16941-215">選択されているアイテムがメッセージであり、**PropertyName** 属性に **BodyAsHTML** を指定する場合は、メッセージを開いて **[ソースの表示]** を選択し、そのアイテムの HTML 表現でのメッセージ本文を確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-215">If the selected item is a message and you specify **BodyAsHTML** in the **PropertyName** attribute, open the message, and then choose **View Source** to verify the message body in the HTML representation of that item.</span></span>

- <span data-ttu-id="16941-216">選択されているアイテムが予定の場合、またはアクティブ化ルールで **PropertyName** に **BodyAsPlaintext** が指定される場合は、Windows での Outlook で Outlook オブジェクト モデルと Visual Basic エディターを使用できます。</span><span class="sxs-lookup"><span data-stu-id="16941-216">If the selected item is an appointment, or if the activation rule specifies **BodyAsPlaintext** in the **PropertyName**, you can use the Outlook object model and the Visual Basic Editor in Outlook on Windows:</span></span>

    1. <span data-ttu-id="16941-217">マクロが有効で、**[開発]** タブが Outlook のリボンに表示されていることを確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-217">Ensure that macros are enabled and the **Developer** tab is displayed in the ribbon for Outlook.</span></span>

    1. <span data-ttu-id="16941-218">Visual Basic エディターで、**[表示]**、**[イミディエイト ウィンドウ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="16941-218">In the Visual Basic Editor, choose **View**, **Immediate Window**.</span></span>

    1. <span data-ttu-id="16941-219">シナリオに応じて各種のプロパティを表示するには、次のように入力します。</span><span class="sxs-lookup"><span data-stu-id="16941-219">Type the following to display various properties depending on the scenario.</span></span>

        - <span data-ttu-id="16941-220">Outlook エクスプローラーで選択されているメッセージ アイテムまたは予定アイテムの HTML 形式の本文。</span><span class="sxs-lookup"><span data-stu-id="16941-220">The HTML body of the message or appointment item selected in the Outlook explorer:</span></span>

        ```vb
        ?ActiveExplorer.Selection.Item(1).HTMLBody
        ```
        - <span data-ttu-id="16941-221">Outlook エクスプローラーで選択されているメッセージ アイテムまたは予定アイテムのプレーン テキスト形式の本文。</span><span class="sxs-lookup"><span data-stu-id="16941-221">The plain text body of the message or appointment item selected in the Outlook explorer:</span></span>

        ```vb
        ?ActiveExplorer.Selection.Item(1).Body
        ```
        - <span data-ttu-id="16941-222">現在の Outlook インスペクターで開かれているメッセージ アイテムまたは予定アイテムの HTML 形式の本文。</span><span class="sxs-lookup"><span data-stu-id="16941-222">The HTML body of the message or appointment item opened in the current Outlook inspector:</span></span>

        ```vb
        ?ActiveInspector.CurrentItem.HTMLBody
        ```
        - <span data-ttu-id="16941-223">現在の Outlook インスペクターで開かれているメッセージ アイテムまたは予定アイテムのプレーン テキスト形式の本文。</span><span class="sxs-lookup"><span data-stu-id="16941-223">The plain text body of the message or appointment item opened in the current Outlook inspector:</span></span>

        ```vb
        ?ActiveInspector.CurrentItem.Body
        ```

<span data-ttu-id="16941-224">**ItemHasRegularExpressionMatch** アクティブ化ルールで **Subject** または **SenderSMTPAddress** が指定される場合、あるいは **ItemIs** ルールまたは **ItemHasAttachment** ルールを使用していて、MAPI の使用に精通しているか使用する必要がある場合は、[MFCMAPI](https://github.com/stephenegriffin/mfcmapi) を使用して、ルールで使用される表 2 の値を確認できます。</span><span class="sxs-lookup"><span data-stu-id="16941-224">If the **ItemHasRegularExpressionMatch** activation rule specifies **Subject** or **SenderSMTPAddress**, or if you use an **ItemIs** or **ItemHasAttachment** rule, and you are familiar with or would like to use MAPI, you can use [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) to verify the value in Table 2 that your rule relies on.</span></span>

<span data-ttu-id="16941-225">**表 2アクティブ化ルールと対応する MAPI プロパティ**</span><span class="sxs-lookup"><span data-stu-id="16941-225">**Table 2. Activation rules and corresponding MAPI properties**</span></span>

|<span data-ttu-id="16941-226">ルールの種類</span><span class="sxs-lookup"><span data-stu-id="16941-226">Type of rule</span></span>|<span data-ttu-id="16941-227">確認する MAPI プロパティ</span><span class="sxs-lookup"><span data-stu-id="16941-227">Verify this MAPI property</span></span>|
|:-----|:-----|
|<span data-ttu-id="16941-228">**ItemHasRegularExpressionMatch** ルールと **Subject**</span><span class="sxs-lookup"><span data-stu-id="16941-228">**ItemHasRegularExpressionMatch** rule with **Subject**</span></span>|[<span data-ttu-id="16941-229">PidTagSubject</span><span class="sxs-lookup"><span data-stu-id="16941-229">PidTagSubject</span></span>](/office/client-developer/outlook/mapi/pidtagsubject-canonical-property)|
|<span data-ttu-id="16941-230">**ItemHasRegularExpressionMatch** ルールと **SenderSMTPAddress**</span><span class="sxs-lookup"><span data-stu-id="16941-230">**ItemHasRegularExpressionMatch** rule with **SenderSMTPAddress**</span></span>|<span data-ttu-id="16941-231">
  [PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) と [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)</span><span class="sxs-lookup"><span data-stu-id="16941-231">[PidTagSenderSmtpAddress](/office/client-developer/outlook/mapi/pidtagsendersmtpaddress-canonical-property) and [PidTagSentRepresentingSmtpAddress](/office/client-developer/outlook/mapi/pidtagsentrepresentingsmtpaddress-canonical-property)</span></span>|
|<span data-ttu-id="16941-232">**ItemIs**</span><span class="sxs-lookup"><span data-stu-id="16941-232">**ItemIs**</span></span>|[<span data-ttu-id="16941-233">PidTagMessageClass</span><span class="sxs-lookup"><span data-stu-id="16941-233">PidTagMessageClass</span></span>](/office/client-developer/outlook/mapi/pidtagmessageclass-canonical-property)|
|<span data-ttu-id="16941-234">**ItemHasAttachment**</span><span class="sxs-lookup"><span data-stu-id="16941-234">**ItemHasAttachment**</span></span>|[<span data-ttu-id="16941-235">PidTagHasAttachments</span><span class="sxs-lookup"><span data-stu-id="16941-235">PidTagHasAttachments</span></span>](/office/client-developer/outlook/mapi/pidtaghasattachments-canonical-property)|

<span data-ttu-id="16941-236">プロパティ値を確認した後、正規表現評価ツールを使用して、正規表現がその値の中で一致を見つけるかどうかをテストできます。</span><span class="sxs-lookup"><span data-stu-id="16941-236">After verifying the property value, you can then use a regular expression evaluation tool to test whether the regular expression finds a match in that value.</span></span>

## <a name="does-the-host-application-apply-all-the-regular-expressions-to-the-portion-of-the-item-body-as-you-expect"></a><span data-ttu-id="16941-237">ホスト アプリケーションはすべての正規表現をアイテムの本文の部分に予期したとおりに適用しますか。</span><span class="sxs-lookup"><span data-stu-id="16941-237">Does the host application apply all the regular expressions to the portion of the item body as you expect?</span></span>

<span data-ttu-id="16941-p123">このセクションは、正規表現を使用するすべてのアクティブ化ルールに適用されます。特に、サイズが大きく、一致の評価に時間がかかるアイテムの本文に適用されます。アクティブ化ルールが依存するアイテム プロパティが予期した値であっても、ホスト アプリケーションがアイテム プロパティの値全体のすべての正規表現を評価できない場合があることに留意してください。適切なパフォーマンスを提供し、閲覧アドイン、Outlook、Outlook on the web、モバイル デバイスによる過度のリソース使用を制御するため、実行時のアクティブ化ルールの正規表現を処理する際には次の制限を確認します。</span><span class="sxs-lookup"><span data-stu-id="16941-p123">This section applies to all activation rules that use regular expressions -- particularly those that are applied to the item body, which may be large in size and take longer to evaluate for matches. You should be aware that even if the item property that an activation rule depends on has the value you expect, the host application may not be able to evaluate all the regular expressions on the entire value of the item property. To provide reasonable performance and to control excessive resource usage by a read add-in, Outlook, Outlook on the web and mobile devices observe the following limits on processing regular expressions in activation rules at run time:</span></span>

- <span data-ttu-id="16941-p124">評価されるアイテムの本文のサイズ -- ホスト アプリケーションが正規表現を評価するアイテムの本文部分には制限があります。これらの制限は、ホスト アプリケーション、フォーム ファクター、およびアイテムの本文の形式によって異なります。詳細については、「[Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」の表 2 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p124">The size of the item body evaluated -- There are limits to the portion of an item body on which the host application evaluates a regular expression. These limits depend on the host application, form factor, and format of the item body. See the details in Table 2 in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).</span></span>

- <span data-ttu-id="16941-p125">正規表現の一致の数 - Outlook リッチ クライアント、Outlook on the web、モバイル デバイスは、それぞれ正規表現の一致を 50 件まで返します。これらの一致は一意であり、重複の一致はこの制限にカウントされません。返される一致の順序を想定しないでください。Outlook リッチ クライアントでの順序は Outlook on the web およびモバイル デバイスでの順序と同じとは限りません。アクティブ化ルールに正規表現の一致が多数存在することが予想されるにもかかわらず、一致が見つからない場合は、この制限を超えている可能性があります。</span><span class="sxs-lookup"><span data-stu-id="16941-p125">Number of regular expression matches -- The Outlook rich clients, and Outlook on the web and mobile devices each returns a maximum of 50 regular expression matches. These matches are unique, and duplicate matches do not count against this limit. Do not assume any order to the returned matches, and do not assume the order in an Outlook rich client is the same as that in Outlook on the web and mobile devices. If you expect many matches to regular expressions in your activation rules, and you're missing a match, you may be exceeding this limit.</span></span>

- <span data-ttu-id="16941-p126">正規表現の一致の長さ -- 正規表現に一致する文字列としてホスト アプリケーションから返される文字列の長さには上限があります。ホスト アプリケーションは上限を超える文字列を一致に含めず、警告メッセージも表示しません。他の regex 評価ツールまたはスタンドアロンの C++ テスト プログラムで正規表現を実行して、このような制限を超える一致があるかどうかを確認できます。表 3 にこの制限の要約を示します。詳細については、「[Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」の表 3 を参照してください。</span><span class="sxs-lookup"><span data-stu-id="16941-p126">Length of a regular expression match -- There are limits to the length of a regular expression match that the host application would return. The host application does not include any match above the limit and does not display any warning message. You can run your regular expression using other regex evaluation tools or a stand-alone C++ test program to verify whether you have a match that exceeds such limits. Table 3 summarizes the limits. For more information, see Table 3 in [Limits for activation and JavaScript API for Outlook add-ins](limits-for-activation-and-javascript-api-for-outlook-add-ins.md).</span></span>

    <span data-ttu-id="16941-253">**表 3正規表現の一致の長さ制限**</span><span class="sxs-lookup"><span data-stu-id="16941-253">**Table 3. Length limits for a regular expression match**</span></span>

    |<span data-ttu-id="16941-254">正規表現の長さ制限</span><span class="sxs-lookup"><span data-stu-id="16941-254">Limit on length of a regex match</span></span>|<span data-ttu-id="16941-255">Outlook リッチ クライアント</span><span class="sxs-lookup"><span data-stu-id="16941-255">Outlook rich clients</span></span>|<span data-ttu-id="16941-256">Outlook on the web またはモバイル デバイス</span><span class="sxs-lookup"><span data-stu-id="16941-256">Outlook on the web or mobile devices</span></span>|
    |:-----|:-----|:-----|
    |<span data-ttu-id="16941-257">アイテムの本文がテキスト形式の場合</span><span class="sxs-lookup"><span data-stu-id="16941-257">Item body is plain text</span></span>|<span data-ttu-id="16941-258">1.5 KB</span><span class="sxs-lookup"><span data-stu-id="16941-258">1.5 KB</span></span>|<span data-ttu-id="16941-259">3 KB</span><span class="sxs-lookup"><span data-stu-id="16941-259">3 KB</span></span>|
    |<span data-ttu-id="16941-260">アイテムの本文が HTML の場合</span><span class="sxs-lookup"><span data-stu-id="16941-260">Item body is HTML</span></span>|<span data-ttu-id="16941-261">3 KB</span><span class="sxs-lookup"><span data-stu-id="16941-261">3 KB</span></span>|<span data-ttu-id="16941-262">3 KB</span><span class="sxs-lookup"><span data-stu-id="16941-262">3 KB</span></span>|

- <span data-ttu-id="16941-p127">Outlook リッチ クライアント用閲覧アドインのすべての正規表現の評価にかかった時間 : 既定では、Outlook はアクティブ化ルール内のすべての正規表現の評価を閲覧アドインごとに 1 秒以内で完了する必要があります。完了しなかった場合、Outlook は最大 3 回まで再試行し、それでも評価を完了できないとアドインを無効化します。Outlook は、アドインが無効になったというメッセージを通知バーに表示します。正規表現に使用可能な時間の長さは、グループ ポリシーまたはレジストリ キーの設定で変更できます。</span><span class="sxs-lookup"><span data-stu-id="16941-p127">Time spent on evaluating all regular expressions of a read add-in for an Outlook rich client: By default, for each read add-in, Outlook must finish evaluating all the regular expressions in its activation rules within 1 second. Otherwise Outlook retries up to three times and disables the add-in if Outlook cannot complete the evaluation. Outlook displays a message in the notification bar that the add-in has been disabled. The amount of time available for your regular expression can be modified by setting a group policy or a registry key.</span></span> 

   > [!NOTE]
   > <span data-ttu-id="16941-267">Outlook リッチ クライアントが、閲覧アドインを無効にした場合、閲覧アドインは、Outlook リッチ クライアント、Outlook on the web、モバイル デバイスの同じメールボックスで使用できなくなります。</span><span class="sxs-lookup"><span data-stu-id="16941-267">If the Outlook rich client disables a read add-in, the read add-in is not available for use for the same mailbox on the Outlook rich client, and Outlook on the web and mobile devices.</span></span>

## <a name="see-also"></a><span data-ttu-id="16941-268">関連項目</span><span class="sxs-lookup"><span data-stu-id="16941-268">See also</span></span>

- [<span data-ttu-id="16941-269">テスト用に Outlook アドインを展開してインストールする</span><span class="sxs-lookup"><span data-stu-id="16941-269">Deploy and install Outlook add-ins for testing</span></span>](testing-and-tips.md)
- [<span data-ttu-id="16941-270">Outlook アドインのアクティブ化ルール</span><span class="sxs-lookup"><span data-stu-id="16941-270">Activation rules for Outlook add-ins</span></span>](activation-rules.md)
- [<span data-ttu-id="16941-271">正規表現アクティブ化ルールを使用して Outlook アドインを表示する</span><span class="sxs-lookup"><span data-stu-id="16941-271">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)
- [<span data-ttu-id="16941-272">Outlook アドインのアクティブ化と JavaScript API の制限</span><span class="sxs-lookup"><span data-stu-id="16941-272">Limits for activation and JavaScript API for Outlook add-ins</span></span>](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [<span data-ttu-id="16941-273">マニフェストの問題を検証し、トラブルシューティングする</span><span class="sxs-lookup"><span data-stu-id="16941-273">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)
