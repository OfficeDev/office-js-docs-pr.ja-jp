---
title: Office 365 管理センターからの一元展開を使用した Office アドインの発行
description: 集中展開を使用して、ISV によって提供される内部アドインとアドインを展開する方法について説明します。
ms.date: 03/22/2021
localization_priority: Normal
ms.openlocfilehash: 3107fc58601683f5356594f2f79ffc5293ea266f
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076652"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a><span data-ttu-id="07618-103">Office 365 管理センターからの一元展開を使用した Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="07618-103">Publish Office Add-ins using Centralized Deployment via the Microsoft 365 admin center</span></span>

<span data-ttu-id="07618-104">このMicrosoft 365 管理センターを使用すると、管理者は組織内のユーザー Officeにアドインを簡単に展開できます。</span><span class="sxs-lookup"><span data-stu-id="07618-104">The Microsoft 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization.</span></span> <span data-ttu-id="07618-105">管理センター経由で展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。</span><span class="sxs-lookup"><span data-stu-id="07618-105">Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required.</span></span> <span data-ttu-id="07618-106">一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="07618-106">You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="07618-107">現在、Microsoft 365 管理センターは次のシナリオをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="07618-107">The Microsoft 365 admin center currently supports the following scenarios.</span></span>

- <span data-ttu-id="07618-108">新しいアドインおよび更新されたアドインの個人、グループ、組織への一元展開。</span><span class="sxs-lookup"><span data-stu-id="07618-108">Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.</span></span>
- <span data-ttu-id="07618-109">複数のクライアント プラットフォームへの展開 (Windows Mac、Web など)。</span><span class="sxs-lookup"><span data-stu-id="07618-109">Deployment to multiple client platforms, including Windows, Mac, and the web.</span></span> <span data-ttu-id="07618-110">このOutlook、iOS と Android への展開もサポートされています。</span><span class="sxs-lookup"><span data-stu-id="07618-110">For Outlook, deployment to iOS and Android is also supported.</span></span> <span data-ttu-id="07618-111">(ただし、iPad での Excel、Outlook、Word、および PowerPoint アドインのユーザー インストールはサポートされていますが、iPad への一元展開 **はサポートされていません**。</span><span class="sxs-lookup"><span data-stu-id="07618-111">(However, while user installation of Excel, Outlook, Word, and PowerPoint add-ins on iPad is supported, Centralized Deployment to iPad is **not** supported.)</span></span>
- <span data-ttu-id="07618-112">英語および世界各国のテナントへの展開。</span><span class="sxs-lookup"><span data-stu-id="07618-112">Deployment to English language and worldwide tenants.</span></span>
- <span data-ttu-id="07618-113">クラウド ホスト型アドインの展開。</span><span class="sxs-lookup"><span data-stu-id="07618-113">Deployment of cloud-hosted add-ins.</span></span>
- <span data-ttu-id="07618-114">ファイアウオール内でホストされているアドインの展開。</span><span class="sxs-lookup"><span data-stu-id="07618-114">Deployment of add-ins that are hosted within a firewall.</span></span>
- <span data-ttu-id="07618-115">AppSource アドインの展開。</span><span class="sxs-lookup"><span data-stu-id="07618-115">Deployment of AppSource add-ins.</span></span>
- <span data-ttu-id="07618-116">ユーザーのアドインの自動インストール (Office アプリケーション起動時)。</span><span class="sxs-lookup"><span data-stu-id="07618-116">Automatic installation of an add-in for users when they launch the Office application.</span></span>
- <span data-ttu-id="07618-117">ユーザーのアドインの自動削除 (管理者がアドインをオフにした場合や削除した場合。または、ユーザーが Azure Active Directory から削除された場合やアドインが展開されているグループから削除された場合)。</span><span class="sxs-lookup"><span data-stu-id="07618-117">Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.</span></span>

<span data-ttu-id="07618-118">組織が集中展開を使用するためのすべての要件を満たしている場合、Microsoft 365 管理者が組織内に Office アドインを展開するための推奨される方法は、集中展開です。</span><span class="sxs-lookup"><span data-stu-id="07618-118">Centralized Deployment is the recommended way for a Microsoft 365 admin to deploy Office Add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment.</span></span> <span data-ttu-id="07618-119">組織で集中展開を使用できるかどうかを判断する方法については、「アドインの集中展開が組織に対して機能するかどうかを判断する」[をMicrosoft 365してください](/office365/admin/manage/centralized-deployment-of-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="07618-119">For information about how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="07618-120">Microsoft 365 への接続がないオンプレミス環境、または Office 2013 を対象とする SharePoint アドインまたは Office アドインを展開するには、SharePoint アプリ カタログを使用[します](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。</span><span class="sxs-lookup"><span data-stu-id="07618-120">In an on-premises environment with no connection to Microsoft 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span> <span data-ttu-id="07618-121">COM/VSTO アドインを展開する場合は、ClickOnce または Windows インストーラーを使用してください。詳細については、「[Office ソリューションの配置](/visualstudio/vsto/deploying-an-office-solution)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="07618-121">To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).</span></span>

## <a name="recommended-approach-for-deploying-office-add-ins"></a><span data-ttu-id="07618-122">Office アドインの展開に推奨されるアプローチ</span><span class="sxs-lookup"><span data-stu-id="07618-122">Recommended approach for deploying Office Add-ins</span></span>

<span data-ttu-id="07618-p105">展開がスムーズに進行するように、段階的なアプローチで Office アドインを展開することを検討してください。以下のプランをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="07618-p105">Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:</span></span>

1. <span data-ttu-id="07618-p106">ビジネス関係者の少人数のグループと IT 部門のメンバーにアドインを展開します。 展開が成功した場合は、第 2 段階に進みます。</span><span class="sxs-lookup"><span data-stu-id="07618-p106">Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.</span></span>

2. <span data-ttu-id="07618-p107">アドインを使用することになるビジネス ユーザーの人数を増やしたグループにアドインを展開します。 展開が成功した場合は、第 3 段階に進みます。</span><span class="sxs-lookup"><span data-stu-id="07618-p107">Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.</span></span>

3. <span data-ttu-id="07618-129">アドインを使用することになるすべてのユーザーのグループにアドインを展開します。</span><span class="sxs-lookup"><span data-stu-id="07618-129">Deploy the add-in to the full set of individuals who will be using the add-in.</span></span>

<span data-ttu-id="07618-130">対象ユーザーの規模に応じて、この手順に段階を追加するか、この手順から段階を削除してください。</span><span class="sxs-lookup"><span data-stu-id="07618-130">Depending on the size of the target audience, you may want to add steps to or remove steps from this procedure.</span></span>

## <a name="publish-an-office-add-in-via-centralized-deployment"></a><span data-ttu-id="07618-131">一元展開による Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="07618-131">Publish an Office Add-in via Centralized Deployment</span></span>

<span data-ttu-id="07618-132">開始する前に、「アドインの集中展開が Microsoft 365 組織で機能するかどうかを判断する」の説明に従って、組織が集中展開の使用に関するすべての要件を満たしていることを[確認](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)します。</span><span class="sxs-lookup"><span data-stu-id="07618-132">Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).</span></span>

<span data-ttu-id="07618-133">組織がすべての要件を満たしている場合は、次に示す手順を実行して、一元展開によって Office アドインを発行します。</span><span class="sxs-lookup"><span data-stu-id="07618-133">If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment:</span></span>

1. <span data-ttu-id="07618-134">仕事または教育アカウントMicrosoft 365にサインインします。</span><span class="sxs-lookup"><span data-stu-id="07618-134">Sign in to Microsoft 365 with your work or education account.</span></span>
1. <span data-ttu-id="07618-135">左上にあるアプリ起動ツールのアイコンを選択して、**[管理]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="07618-135">Select the app launcher icon in the upper-left and choose **Admin**.</span></span>
1. <span data-ttu-id="07618-136">ナビゲーション メニューで、[詳細の表示]**を** 選択し、[統合  >  **設定] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="07618-136">In the navigation menu, select **Show more**, then choose **Settings** > **Integrated apps**.</span></span>
1. <span data-ttu-id="07618-137">ページの上部で、[アドイン] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="07618-137">At the top of the page, choose **Add-ins**.</span></span>
1. <span data-ttu-id="07618-138">新しい Microsoft 365 管理センター を発表するメッセージがページの上部に表示される場合は、管理センター プレビューに移動するメッセージを選択します (「概要」を[参照](/microsoft-365/admin/admin-overview/about-the-admin-center)Microsoft 365 管理センター)。</span><span class="sxs-lookup"><span data-stu-id="07618-138">If you see a message on the top of the page announcing the new Microsoft 365 admin center, choose the message to go to the Admin Center Preview (see [About the Microsoft 365 admin center](/microsoft-365/admin/admin-overview/about-the-admin-center)).</span></span>
1. <span data-ttu-id="07618-139">ページの上部にある **[アドインの展開]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-139">Choose **Deploy Add-In** at the top of the page.</span></span>
1. <span data-ttu-id="07618-140">要件の確認後、**[次へ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-140">Choose **Next** after reviewing the requirements.</span></span>
1. <span data-ttu-id="07618-141">**[一元展開]** ページで、次のいずれかのオプションを選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-141">Choose one of the following options on the **Centralized Deployment** page:</span></span>

    - <span data-ttu-id="07618-142">**Office ストアからアドインを追加します。**</span><span class="sxs-lookup"><span data-stu-id="07618-142">**I want to add an Add-In from the Office Store.**</span></span>
    - <span data-ttu-id="07618-p108">**このデバイスにマニフェスト ファイル (.xml) があります。** このオプションの場合は、**[参照]** を選択して、使用するマニフェスト ファイル (.xml) の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="07618-p108">**I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.</span></span>
    - <span data-ttu-id="07618-p109">**マニフェスト ファイルの URL がわかります。** このオプションの場合は、所定のフィールドにマニフェストの URL を入力します。</span><span class="sxs-lookup"><span data-stu-id="07618-p109">**I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.</span></span>

    ![[Add-In] ダイアログ ボックスMicrosoft 365 管理センター。](../images/new-add-in.png)

8. <span data-ttu-id="07618-148">Office ストアからアドインを追加するオプションを選択した場合は、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-148">If you selected the option to add an add-in from the Office Store, select the add-in.</span></span> <span data-ttu-id="07618-149">選択可能なアドインは、**[あなたへのおすすめ]**、**[評価]**、**[名前]** のカテゴリから表示できます。</span><span class="sxs-lookup"><span data-stu-id="07618-149">You can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**.</span></span> <span data-ttu-id="07618-150">Office ストアからは無料のアドインのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="07618-150">You may only add free add-ins from Office Store.</span></span> <span data-ttu-id="07618-151">有料のアドインの追加は、現在はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="07618-151">Adding paid add-ins isn't currently supported.</span></span>

    > [!NOTE]
    > <span data-ttu-id="07618-152">Office ストアのオプションでは、管理者の操作なしで、ユーザーが自動的にアドインの更新と機能強化を利用できます。</span><span class="sxs-lookup"><span data-stu-id="07618-152">With the Office Store option, updates and enhancements to the add-in are automatically available to users without your intervention.</span></span>

    ![[アドイン] ダイアログボックスを選択Microsoft 365 管理センター。](../images/select-an-add-in.png)

9. <span data-ttu-id="07618-154">アドイン **の詳細** 、プライバシー ポリシー、ライセンス条項を確認した後、[続行] を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-154">Choose **Continue** after reviewing the add-in details, Privacy Policy, and License Terms.</span></span>

    ![[アドイン] ページで選択Microsoft 365 管理センター。](../images/selected-add-in-admin-center.png)

10. <span data-ttu-id="07618-156">[ユーザーの **割り当て] ページ** で、[ **すべてのユーザー**] 、[ **特定のユーザー/グループ]**、または [自分のみ] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="07618-156">On the **Assign Users** page, choose **Everyone**, **Specific Users/Groups**, or **Only me**.</span></span> <span data-ttu-id="07618-157">検索ボックスを使用して、アドインを展開するユーザーやグループを検索します。</span><span class="sxs-lookup"><span data-stu-id="07618-157">Use the search box to find the users and groups to whom you want to deploy the add-in.</span></span> <span data-ttu-id="07618-158">アドインOutlook、展開方法 [固定] 、[利用可能]、または **[オプション**] を選択 **できます**。</span><span class="sxs-lookup"><span data-stu-id="07618-158">For Outlook add-ins, you can also choose the deployment method **Fixed**, **Available**, or **Optional**.</span></span>

    ![アクセスおよび展開方法を持つユーザーを管理Microsoft 365 管理センター。](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > <span data-ttu-id="07618-160">シングル サインオン [(SSO)](../develop/sso-in-office-add-ins.md) を使用するアドインは、アドイン マニフェストに記載されているスコープに同意するよう管理者に求めるメッセージを表示します。</span><span class="sxs-lookup"><span data-stu-id="07618-160">Add-ins that utilize [single sign-on (SSO)](../develop/sso-in-office-add-ins.md) will prompt the admin to consent to the scopes listed in the add-in manifest.</span></span>  <span data-ttu-id="07618-161">同じバッキング サービスが複数のアドインで使用されている場合 (同じ Azure App ID が異なるアドインの SSO で使用される)、各アドインのスコープに対して、各展開に対する同意を求めるメッセージが表示されます。</span><span class="sxs-lookup"><span data-stu-id="07618-161">If the same backing service is used across multiple add-ins (the same Azure App ID is used with SSO in different add-ins), the scopes for each add-in will be prompted for consent with each deployment.</span></span> <span data-ttu-id="07618-162">このページには、アドインで必要なアクセス許可の一覧も表示されます。</span><span class="sxs-lookup"><span data-stu-id="07618-162">This page will also display the list of permissions that the add-in requires.</span></span>

11. <span data-ttu-id="07618-163">完了したら、[展開] **を選択します**。</span><span class="sxs-lookup"><span data-stu-id="07618-163">When finished, choose **Deploy**.</span></span> <span data-ttu-id="07618-164">このプロセスには、最大で 3 分かかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="07618-164">This process may take up to three minutes.</span></span> <span data-ttu-id="07618-165">その後、**[次へ]** を押してチュートリアルを終了します。</span><span class="sxs-lookup"><span data-stu-id="07618-165">Then, finish the walkthrough by pressing **Next**.</span></span> <span data-ttu-id="07618-166">これで、他のアプリと共にアドインがOfficeされます。</span><span class="sxs-lookup"><span data-stu-id="07618-166">You now see your add-in along with other Office apps.</span></span>

    > [!NOTE]
    > <span data-ttu-id="07618-167">管理者が [展開] を **選択すると**、すべてのユーザーに同意が与えられる。</span><span class="sxs-lookup"><span data-stu-id="07618-167">When an administrator chooses **Deploy**, consent is given for all users.</span></span>

    ![アプリの一覧 (Microsoft 365 管理センター)。](../images/citations.png)

> [!TIP]
> <span data-ttu-id="07618-169">組織内のユーザーやグループに新しいアドインを展開するときには、いつどのようにアドインを使用するかについての説明と、サポート資料 (関連するヘルプ コンテンツやよくある質問など) へのリンクを含む電子メールの送信を検討してください。</span><span class="sxs-lookup"><span data-stu-id="07618-169">When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.</span></span>

## <a name="considerations-when-granting-access-to-an-add-in"></a><span data-ttu-id="07618-170">アドインへのアクセスを許可するときの考慮事項</span><span class="sxs-lookup"><span data-stu-id="07618-170">Considerations when granting access to an add-in</span></span>

<span data-ttu-id="07618-p114">管理者は、組織内のすべてのユーザーにアドインを割り当てることも、組織内の特定のユーザーやグループにアドインを割り当てることもできます。 次のリストに、それぞれのオプションの影響を示します。</span><span class="sxs-lookup"><span data-stu-id="07618-p114">Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:</span></span>

- <span data-ttu-id="07618-p115">**すべてのユーザー**: 名前が示すように、このオプションでは、テナント内のすべてのユーザーにアドインが割り当てられます。対象のアドインが組織全体で汎用な場合にのみ、このオプションを慎重に使用します。</span><span class="sxs-lookup"><span data-stu-id="07618-p115">**Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.</span></span>

- <span data-ttu-id="07618-p116">**ユーザー**: 個別のユーザーにアドインを割り当てる場合は、追加のユーザーにアドインを割り当てるたびに、アドインの一元展開の設定を更新する必要があります。 同様に、ユーザーのアドインへのアクセス権を削除するたびに、アドインの一元展開の設定を更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07618-p116">**Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.</span></span>

- <span data-ttu-id="07618-177">**グループ**: グループにアドインを割り当てると、グループに追加されたユーザーにアドインが自動的に割り当てられます。</span><span class="sxs-lookup"><span data-stu-id="07618-177">**Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in.</span></span> <span data-ttu-id="07618-178">同様に、ユーザーがグループから削除されると、そのユーザーはアドインへのアクセス権を自動的に失います。</span><span class="sxs-lookup"><span data-stu-id="07618-178">Likewise, when a user is removed from a group, the user automatically loses access to the add-in.</span></span> <span data-ttu-id="07618-179">いずれの場合も、管理者から追加のアクションはMicrosoft 365ありません。</span><span class="sxs-lookup"><span data-stu-id="07618-179">In either case, no additional action is required from the Microsoft 365 admin.</span></span>

<span data-ttu-id="07618-p118">一般に、保守が簡単になるように、可能な場合は常にグループを使用してアドインを割り当てるようにしてください。 ただし、アドインのアクセスをユーザーの非常に少数のメンバーに制限する場合は、具体的なユーザーにアドインを割り当てるほうが実用的です。</span><span class="sxs-lookup"><span data-stu-id="07618-p118">In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.</span></span>

## <a name="add-in-states"></a><span data-ttu-id="07618-182">アドインの状態</span><span class="sxs-lookup"><span data-stu-id="07618-182">Add-in states</span></span>

<span data-ttu-id="07618-183">次の表では、様々なアドインの状態について説明しています。</span><span class="sxs-lookup"><span data-stu-id="07618-183">The following table describes the different states of an add-in.</span></span>

|<span data-ttu-id="07618-184">状態</span><span class="sxs-lookup"><span data-stu-id="07618-184">State</span></span>|<span data-ttu-id="07618-185">状態が発生する原因</span><span class="sxs-lookup"><span data-stu-id="07618-185">How the state occurs</span></span>|<span data-ttu-id="07618-186">影響</span><span class="sxs-lookup"><span data-stu-id="07618-186">Impact</span></span>|
|-----|--------------------|------|
|<span data-ttu-id="07618-187">**アクティブ**</span><span class="sxs-lookup"><span data-stu-id="07618-187">**Active**</span></span>|<span data-ttu-id="07618-188">管理者がアドインをアップロードして、ユーザーやグループに割り当てた。</span><span class="sxs-lookup"><span data-stu-id="07618-188">Admin uploaded the add-in and assigned it to users and/or groups.</span></span>|<span data-ttu-id="07618-189">アドインが割り当てられたユーザーやグループは、関連する Office クライアントでアドインを表示できます。</span><span class="sxs-lookup"><span data-stu-id="07618-189">Users and/or groups assigned to the add-in see it in the relevant Office clients.</span></span>|
|<span data-ttu-id="07618-190">**オフ**</span><span class="sxs-lookup"><span data-stu-id="07618-190">**Turned off**</span></span>|<span data-ttu-id="07618-191">管理者がアドインをオフにした。</span><span class="sxs-lookup"><span data-stu-id="07618-191">Admin turned off the add-in.</span></span>|<span data-ttu-id="07618-p119">アドインに割り当てられたユーザーやグループは、アドインにアクセスできなくなります。 アドインの状態が **[オフ]** から **[アクティブ]** に変更されると、ユーザーやグループはアドインに再度アクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="07618-p119">Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.</span></span>|
|<span data-ttu-id="07618-194">**Deleted**</span><span class="sxs-lookup"><span data-stu-id="07618-194">**Deleted**</span></span>|<span data-ttu-id="07618-195">管理者がアドインを削除した。</span><span class="sxs-lookup"><span data-stu-id="07618-195">Admin deleted the add-in.</span></span>|<span data-ttu-id="07618-196">アドインが割り当てられたユーザーやグループは、そのアドインにアクセスできなくなります。</span><span class="sxs-lookup"><span data-stu-id="07618-196">Users and/or groups assigned the add-in no longer have access to it.</span></span>|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a><span data-ttu-id="07618-197">一元展開によって発行された Office アドインの更新</span><span class="sxs-lookup"><span data-stu-id="07618-197">Updating Office Add-ins that are published via Centralized Deployment</span></span>

<span data-ttu-id="07618-198">Officeを一元的な展開で公開すると、その変更が Web アプリケーションに実装された後、アドインの Web アプリケーションに加えた変更は自動的にすべてのユーザーが利用できます。</span><span class="sxs-lookup"><span data-stu-id="07618-198">After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users after those changes are implemented in the web application.</span></span> <span data-ttu-id="07618-199">たとえば、アドインのアイコン、テキスト、アドイン のコマンドを更新するためにアドインの [XML](../develop/add-in-manifests.md) マニフェスト ファイルに加えた変更は、次のように行われます。</span><span class="sxs-lookup"><span data-stu-id="07618-199">Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md) to, for example, update the add-in's icon, text, or add-in commands, happen as follows:</span></span>

- <span data-ttu-id="07618-200">**Line-of-business** アドイン : Microsoft 365 管理センター 経由で集中展開を実装する際に、管理者がマニフェスト ファイルを明示的にアップロードした場合 (デバイスまたは URL を指す) 場合、管理者は目的の変更を含む新しいマニフェスト ファイルをアップロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="07618-200">**Line-of-business add-in**: If an admin explicitly uploaded a manifest file (either from their device or by pointing to a URL) when implementing Centralized Deployment via the Microsoft 365 admin center, the admin must upload a new manifest file that contains the desired changes.</span></span> <span data-ttu-id="07618-201">更新したマニフェスト ファイルがアップロードされると、関連する Office アプリケーションの次回起動時にアドインが更新されます。</span><span class="sxs-lookup"><span data-stu-id="07618-201">After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.</span></span>

  > [!NOTE]
  > <span data-ttu-id="07618-202">管理者は、更新を行う LOB アドインを削除する必要があります。</span><span class="sxs-lookup"><span data-stu-id="07618-202">An admin doesn't need to remove a LOB add-in to make an update.</span></span> <span data-ttu-id="07618-203">[アドイン] セクションでは、管理者は LOB アドインを選択するだけで、右下隅にある [アドインの更新]ボタンを押してこの機能を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="07618-203">In the Add-ins section, the admin can simply choose the LOB add-in and invoke this functionality by pressing the **Update add-in** button present in the bottom right corner.</span></span>
  >
  > ![スクリーンショットは、[アドインの更新] ダイアログを表示Microsoft 365 管理センター。](../images/update-add-in-admin-center.png)

- <span data-ttu-id="07618-205">**Office** ストア アドイン : Microsoft 365 管理センター 経由で集中展開を実装するときに管理者が Office ストアからアドインを選択し、Office ストアでアドインの更新を行った場合、アドインは後で集中展開を介して更新されます。</span><span class="sxs-lookup"><span data-stu-id="07618-205">**Office Store add-in**: If an admin selected an add-in from the Office Store when implementing Centralized Deployment via the Microsoft 365 admin center, and the add-in updates in the Office Store, the add-in will update later via Centralized Deployment.</span></span> <span data-ttu-id="07618-206">すべてのエンド ユーザーに対してストア アドインの更新プログラムがフローするには、最大 24 時間かかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="07618-206">It can take up to 24 hours for the Store add-in updates to flow for all end users.</span></span> <span data-ttu-id="07618-207">この期間が経過すると、Officeアプリケーションが再起動すると、アドインが更新されます。</span><span class="sxs-lookup"><span data-stu-id="07618-207">After this duration, the next time the relevant Office applications restart for these users, the add-in will update.</span></span> <span data-ttu-id="07618-208">ユーザーは、[タブ アドインの挿入] [管理] タブの [ヒット更新] を選択して、手動更新をトリガーして最新のストア アドイン バージョン  >    >    >  **を取得することもできます**。</span><span class="sxs-lookup"><span data-stu-id="07618-208">Users can also trigger a Manual Refresh to get the latest Store add-in version by selecting **Insert Tab** > **Add-ins** > **Admin Managed Tab** > **Hit Refresh**.</span></span>

## <a name="end-user-experience-with-add-ins"></a><span data-ttu-id="07618-209">アドインのエンド ユーザー エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="07618-209">End user experience with add-ins</span></span>

<span data-ttu-id="07618-210">一元展開によるアドインの発行が完了すると、エンド ユーザーはアドインがサポートする任意のプラットフォームでアドインの使用を開始できます。</span><span class="sxs-lookup"><span data-stu-id="07618-210">After an add-in has been published via Centralized Deployment, end users may start using it on any platform that the add-in supports.</span></span>

<span data-ttu-id="07618-p124">アドインでアドイン コマンドがサポートされている場合、アドインが展開されているすべてのユーザーに対して、コマンドが Office アプリケーション リボンに表示されます。 次の例では、**[引用文献]** アドインのリボンに **[引用文献の検索]** コマンドが表示されています。</span><span class="sxs-lookup"><span data-stu-id="07618-p124">If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.</span></span>

![スクリーンショットは、[引用文献] アドインOffice アプリ[引用文献の検索] コマンドが強調表示されたリボンのセクションを示しています。](../images/search-citation.png)

<span data-ttu-id="07618-214">アドインがアドイン コマンドをサポートしていない場合、ユーザーは次の手順を実行することで、Office アプリケーションにアドインを追加できます。</span><span class="sxs-lookup"><span data-stu-id="07618-214">If the add-in does not support add-in commands, users can add it to their Office application by doing the following:</span></span>

1. <span data-ttu-id="07618-215">Word 2016 以降、Excel 2016 以降、または PowerPoint 2016 以降で **[挿入]** > **[個人用アドイン]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-215">In Word 2016 or later, Excel 2016 or later, or PowerPoint 2016 or later, choose **Insert** > **My Add-ins**.</span></span>
2. <span data-ttu-id="07618-216">アドイン ウィンドウで **[管理者が管理]** タブを選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-216">Choose the **Admin Managed** tab in the add-in window.</span></span>
3. <span data-ttu-id="07618-217">アドインを選択して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-217">Choose the add-in, and then choose **Add**.</span></span>

    ![Office アプリケーションの [Office アドイン] ページにある [管理者が管理] タブを示すスクリーンショット。 [引用文献] アドインがタブに表示されます。](../images/office-add-ins-admin-managed.png)

<span data-ttu-id="07618-220">ただし、Outlook 2016 以降では、ユーザーは次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="07618-220">However, for Outlook 2016 or later, users can do the following:</span></span>

1. <span data-ttu-id="07618-221">Outlook で **[ホーム]** > **[ストア]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-221">In Outlook, choose **Home** > **Store**.</span></span>
2. <span data-ttu-id="07618-222">アドイン タブの下にある **[管理者が管理]** の項目を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-222">Choose the **Admin-managed** item under the add-in tab.</span></span>
3. <span data-ttu-id="07618-223">アドインを選択して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="07618-223">Choose the add-in, and then choose **Add**.</span></span>

    ![Outlook アプリケーションの [ストア] ページの [管理者が管理] 領域を示すスクリーンショット。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a><span data-ttu-id="07618-225">関連項目</span><span class="sxs-lookup"><span data-stu-id="07618-225">See also</span></span>

- [<span data-ttu-id="07618-226">アドインの一元展開が Microsoft 365 組織で動作するかどうかを判断する</span><span class="sxs-lookup"><span data-stu-id="07618-226">Determine if Centralized Deployment of add-ins works for your Microsoft 365 organization</span></span>](/office365/admin/manage/centralized-deployment-of-add-ins)
