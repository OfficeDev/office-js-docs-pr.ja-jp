---
title: Office 365 管理センターからの一元展開を使用した Office アドインの発行
description: 一元展開を使用して、内部アドインと Isv が提供するアドインを展開する方法について説明します。
ms.date: 03/24/2020
localization_priority: Normal
ms.openlocfilehash: 4c19a272e448e38bb5e895cd0bc2a53707a172ad
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217775"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-office-365-admin-center"></a><span data-ttu-id="7bdc3-103">Office 365 管理センターからの一元展開を使用した Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="7bdc3-103">Publish Office Add-ins using Centralized Deployment via the Office 365 admin center</span></span>

<span data-ttu-id="7bdc3-p101">Office 365 管理センターを使用すると、管理者は組織内のユーザーやグループに簡単に Office アドインを展開できます。管理センター経由で展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p101">The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.</span></span>

<span data-ttu-id="7bdc3-107">現在、Office 365 管理センターは次のシナリオをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-107">The Office 365 admin center currently supports the following scenarios:</span></span>

- <span data-ttu-id="7bdc3-108">新しいアドインおよび更新されたアドインの個人、グループ、組織への一元展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-108">Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.</span></span>
- <span data-ttu-id="7bdc3-109">Windows、Mac、iOS、Android、web 上の複数のプラットフォームへの展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-109">Deployment to multiple platforms, including Windows, Mac, iOS, Android, and on the web.</span></span>
- <span data-ttu-id="7bdc3-110">英語および世界各国のテナントへの展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-110">Deployment to English language and worldwide tenants.</span></span>
- <span data-ttu-id="7bdc3-111">クラウド ホスト型アドインの展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-111">Deployment of cloud-hosted add-ins.</span></span>
- <span data-ttu-id="7bdc3-112">ファイアウオール内でホストされているアドインの展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-112">Deployment of add-ins that are hosted within a firewall.</span></span>
- <span data-ttu-id="7bdc3-113">AppSource アドインの展開。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-113">Deployment of AppSource add-ins.</span></span>
- <span data-ttu-id="7bdc3-114">ユーザーのアドインの自動インストール (Office アプリケーション起動時)。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-114">Automatic installation of an add-in for users when they launch the Office application.</span></span>
- <span data-ttu-id="7bdc3-115">ユーザーのアドインの自動削除 (管理者がアドインをオフにした場合や削除した場合。または、ユーザーが Azure Active Directory から削除された場合やアドインが展開されているグループから削除された場合)。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-115">Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.</span></span>

<span data-ttu-id="7bdc3-116">一元展開は、一元展開を使用するためのすべての要件を組織が満たしているときに、Office 365 管理者が組織内で Office アドインを展開する場合に推奨される方法です。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-116">Centralized Deployment is the recommended way for an Office 365 admin to deploy Office Add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment.</span></span> <span data-ttu-id="7bdc3-117">組織で一元展開を使用できるかどうかを判断する方法の詳細については、「[アドインの一元展開が Office 365 組織で動作するかどうかを判断する](/office365/admin/manage/centralized-deployment-of-add-ins)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-117">For information about how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Office 365 organization](/office365/admin/manage/centralized-deployment-of-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="7bdc3-p103">Office 365 に接続していないオンプレミス環境の場合や、Office 2013 を対象とした SharePoint アドインまたは Office アドインを展開する場合は、[SharePoint アプリ カタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)を使用してください。COM/VSTO アドインを展開する場合は、ClickOnce または Windows インストーラーを使用してください。詳細については、「[Office ソリューションの配置](/visualstudio/vsto/deploying-an-office-solution)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p103">In an on-premises environment with no connection to Office 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint app catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](/visualstudio/vsto/deploying-an-office-solution).</span></span>

## <a name="recommended-approach-for-deploying-office-add-ins"></a><span data-ttu-id="7bdc3-120">Office アドインの展開に推奨されるアプローチ</span><span class="sxs-lookup"><span data-stu-id="7bdc3-120">Recommended approach for deploying Office Add-ins</span></span>

<span data-ttu-id="7bdc3-p104">展開がスムーズに進行するように、段階的なアプローチで Office アドインを展開することを検討してください。以下のプランをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p104">Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:</span></span>

1. <span data-ttu-id="7bdc3-p105">ビジネス関係者の少人数のグループと IT 部門のメンバーにアドインを展開します。 展開が成功した場合は、第 2 段階に進みます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p105">Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.</span></span>

2. <span data-ttu-id="7bdc3-p106">アドインを使用することになるビジネス ユーザーの人数を増やしたグループにアドインを展開します。 展開が成功した場合は、第 3 段階に進みます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p106">Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.</span></span>

3. <span data-ttu-id="7bdc3-127">アドインを使用することになるすべてのユーザーのグループにアドインを展開します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-127">Deploy the add-in to the full set of individuals who will be using the add-in.</span></span>

<span data-ttu-id="7bdc3-128">対象ユーザーの規模に応じて、この手順に段階を追加するか、この手順から段階を削除してください。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-128">Depending on the size of the target audience, you may want to add steps to or remove steps from this procedure.</span></span>

## <a name="publish-an-office-add-in-via-centralized-deployment"></a><span data-ttu-id="7bdc3-129">一元展開による Office アドインの発行</span><span class="sxs-lookup"><span data-stu-id="7bdc3-129">Publish an Office Add-in via Centralized Deployment</span></span>

<span data-ttu-id="7bdc3-130">作業を開始する前に、組織が一元展開を使用するためのすべての要件を満たしていることを確認してください。詳細については、「[アドインの一元展開が Office 365 組織で動作するかどうかを判断する](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)」を参照してください。 </span><span class="sxs-lookup"><span data-stu-id="7bdc3-130">Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Office 365 organization](/microsoft-365/admin/manage/centralized-deployment-of-add-ins).</span></span>

<span data-ttu-id="7bdc3-131">組織がすべての要件を満たしている場合は、次に示す手順を実行して、一元展開によって Office アドインを発行します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-131">If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment:</span></span>

1. <span data-ttu-id="7bdc3-132">職場または学校のアカウントを使用して、Office 365 にサインインします。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-132">Sign in to Office 365 with your work or school account.</span></span>
2. <span data-ttu-id="7bdc3-133">左上にあるアプリ起動ツールのアイコンを選択して、**[管理]** をクリックします。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-133">Select the app launcher icon in the upper-left and choose **Admin**.</span></span>
3. <span data-ttu-id="7bdc3-134">ナビゲーション メニューで、**[表示数を増やす]** を押し、**[設定]** > **[サービスとアドイン]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-134">In the navigation menu, press **Show more**, then choose **Settings** > **Services & add-ins**.</span></span>
4. <span data-ttu-id="7bdc3-135">ページの上部に新しい Office 365 管理センターについて通知するメッセージが表示されている場合は、そのメッセージを選択して管理センター プレビューに移動します (「[Office 365 管理センターについて](/microsoft-365/admin/admin-overview/about-the-admin-center)」を参照)。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-135">If you see a message on the top of the page announcing the new Office 365 admin center, choose the message to go to the Admin Center Preview (see [About the Office 365 admin center](/microsoft-365/admin/admin-overview/about-the-admin-center)).</span></span>
5. <span data-ttu-id="7bdc3-136">ページの上部にある **[アドインの展開]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-136">Choose **Deploy Add-In** at the top of the page.</span></span>
6. <span data-ttu-id="7bdc3-137">要件の確認後、**[次へ]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-137">Choose **Next** after reviewing the requirements.</span></span>
7. <span data-ttu-id="7bdc3-138">**[一元展開]** ページで、次のいずれかのオプションを選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-138">Choose one of the following options on the **Centralized Deployment** page:</span></span>

    - <span data-ttu-id="7bdc3-139">**Office ストアからアドインを追加します。**</span><span class="sxs-lookup"><span data-stu-id="7bdc3-139">**I want to add an Add-In from the Office Store.**</span></span>
    - <span data-ttu-id="7bdc3-p107">**このデバイスにマニフェスト ファイル (.xml) があります。** このオプションの場合は、**[参照]** を選択して、使用するマニフェスト ファイル (.xml) の場所を指定します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p107">**I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.</span></span>
    - <span data-ttu-id="7bdc3-p108">**マニフェスト ファイルの URL がわかります。** このオプションの場合は、所定のフィールドにマニフェストの URL を入力します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p108">**I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.</span></span>

    ![Office 365 管理センターの [新しいアドイン] ダイアログ](../images/new-add-in.png)

8. <span data-ttu-id="7bdc3-145">Office ストアからアドインを追加するオプションを選択した場合は、アドインを選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-145">If you selected the option to add an add-in from the Office Store, select the add-in.</span></span> <span data-ttu-id="7bdc3-146">選択可能なアドインは、**[あなたへのおすすめ]**、**[評価]**、**[名前]** のカテゴリから表示できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-146">You can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**.</span></span> <span data-ttu-id="7bdc3-147">Office ストアからは無料のアドインのみを追加できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-147">You may only add free add-ins from Office Store.</span></span> <span data-ttu-id="7bdc3-148">有料のアドインの追加は、現在はサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-148">Adding paid add-ins isn't currently supported.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7bdc3-149">Office ストアのオプションでは、管理者の操作なしで、ユーザーが自動的にアドインの更新と機能強化を利用できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-149">With the Office Store option, updates and enhancements to the add-in are automatically available to users without your intervention.</span></span>

    ![Office 365 管理センターでアドインダイアログを選択する](../images/select-an-add-in.png)

9. <span data-ttu-id="7bdc3-151">アドインの詳細、プライバシーポリシー、ライセンス条項を確認した後、[**続行**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-151">Choose **Continue** after reviewing the add-in details, Privacy Policy, and License Terms.</span></span>

    ![Office 365 管理センターで選択されているアドインページ](../images/selected-add-in-admin-center.png)

10. <span data-ttu-id="7bdc3-153">[**ユーザーの割り当て**] ページで、[**すべて**のユーザー]、[**特定のユーザー/グループ**]、または [**自分のみ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-153">On the **Assign Users** page, choose **Everyone**, **Specific Users/Groups**, or **Only me**.</span></span> <span data-ttu-id="7bdc3-154">検索ボックスを使用して、アドインを展開するユーザーやグループを検索します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-154">Use the search box to find the users and groups to whom you want to deploy the add-in.</span></span> <span data-ttu-id="7bdc3-155">Outlook アドインの場合は、展開方法として**Fixed**、 **Available**、または**Optional**を選択することもできます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-155">For Outlook add-ins, you can also choose the deployment method **Fixed**, **Available**, or **Optional**.</span></span>

    ![Office 365 管理センターでアクセスと展開の方法を管理する](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > <span data-ttu-id="7bdc3-157">アドイン用の[シングル サインオン (SSO)](../develop/sso-in-office-add-ins.md) システムは現在プレビューなので、運用環境のアドインとして使用してはいけません。SSO を使用するアドインが展開されている場合、割り当てられているユーザーとグループは、同じ Azure アプリ ID を共有するアドインによっても共有されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-157">A [single sign-on (SSO)](../develop/sso-in-office-add-ins.md) system for add-ins is currently in preview and should not be used for production add-ins. When an add-in using SSO is deployed, the users and groups assigned are also shared with add-ins that share the same Azure App ID.</span></span> <span data-ttu-id="7bdc3-158">ユーザーの割り当ての変更は、これらのアドインにも適用されます。関連するアドインは、このページに表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-158">Any changes to user assignments are also applied to those add-ins. The related add-ins are shown on this page.</span></span> <span data-ttu-id="7bdc3-159">SSO アドインに限り、アドインで必要な Microsoft Graph アクセス許可のリストがこのページに表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-159">For SSO add-ins only, this page displays the list of Microsoft Graph permissions that the add-in requires.</span></span>

11. <span data-ttu-id="7bdc3-160">完了したら、[**展開**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-160">When finished, choose **Deploy**.</span></span> <span data-ttu-id="7bdc3-161">このプロセスには、最大で 3 分かかる場合があります。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-161">This process may take up to three minutes.</span></span> <span data-ttu-id="7bdc3-162">その後、**[次へ]** を押してチュートリアルを終了します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-162">Then, finish the walkthrough by pressing **Next**.</span></span> <span data-ttu-id="7bdc3-163">アドインが Office 365 のその他のアプリと共に表示されるようになります。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-163">You now see your add-in along with other apps in Office 365.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7bdc3-164">管理者が [**展開**] を選択すると、すべてのユーザーに同意が与えられます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-164">When an administrator chooses **Deploy**, consent is given for all users.</span></span>

    ![Office 365 管理センターのアプリのリスト](../images/citations.png)

> [!TIP]
> <span data-ttu-id="7bdc3-166">組織内のユーザーやグループに新しいアドインを展開するときには、いつどのようにアドインを使用するかについての説明と、サポート資料 (関連するヘルプ コンテンツやよくある質問など) へのリンクを含む電子メールの送信を検討してください。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-166">When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.</span></span>

## <a name="considerations-when-granting-access-to-an-add-in"></a><span data-ttu-id="7bdc3-167">アドインへのアクセスを許可するときの考慮事項</span><span class="sxs-lookup"><span data-stu-id="7bdc3-167">Considerations when granting access to an add-in</span></span>

<span data-ttu-id="7bdc3-p113">管理者は、組織内のすべてのユーザーにアドインを割り当てることも、組織内の特定のユーザーやグループにアドインを割り当てることもできます。 次のリストに、それぞれのオプションの影響を示します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p113">Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:</span></span>

- <span data-ttu-id="7bdc3-p114">**すべてのユーザー**: 名前が示すように、このオプションでは、テナント内のすべてのユーザーにアドインが割り当てられます。対象のアドインが組織全体で汎用な場合にのみ、このオプションを慎重に使用します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p114">**Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.</span></span>

- <span data-ttu-id="7bdc3-p115">**ユーザー**: 個別のユーザーにアドインを割り当てる場合は、追加のユーザーにアドインを割り当てるたびに、アドインの一元展開の設定を更新する必要があります。 同様に、ユーザーのアドインへのアクセス権を削除するたびに、アドインの一元展開の設定を更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p115">**Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.</span></span>

- <span data-ttu-id="7bdc3-p116">**グループ**: グループにアドインを割り当てると、グループに追加されたユーザーにアドインが自動的に割り当てられます。 同様に、ユーザーがグループから削除されると、そのユーザーはアドインへのアクセス権を自動的に失います。 どちらの場合も、Office 365 管理から追加の操作を実行する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p116">**Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in. Likewise, when a user is removed from a group, the user automatically loses access to the add-in. In either case, no additional action is required from the Office 365 admin.</span></span>

<span data-ttu-id="7bdc3-p117">一般に、保守が簡単になるように、可能な場合は常にグループを使用してアドインを割り当てるようにしてください。 ただし、アドインのアクセスをユーザーの非常に少数のメンバーに制限する場合は、具体的なユーザーにアドインを割り当てるほうが実用的です。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p117">In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.</span></span>

## <a name="add-in-states"></a><span data-ttu-id="7bdc3-179">アドインの状態</span><span class="sxs-lookup"><span data-stu-id="7bdc3-179">Add-in states</span></span>

<span data-ttu-id="7bdc3-180">次の表では、様々なアドインの状態について説明しています。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-180">The following table describes the different states of an add-in.</span></span>

|<span data-ttu-id="7bdc3-181">状態</span><span class="sxs-lookup"><span data-stu-id="7bdc3-181">State</span></span>|<span data-ttu-id="7bdc3-182">状態が発生する原因</span><span class="sxs-lookup"><span data-stu-id="7bdc3-182">How the state occurs</span></span>|<span data-ttu-id="7bdc3-183">影響</span><span class="sxs-lookup"><span data-stu-id="7bdc3-183">Impact</span></span>|
|-----|--------------------|------|
|<span data-ttu-id="7bdc3-184">**アクティブ**</span><span class="sxs-lookup"><span data-stu-id="7bdc3-184">**Active**</span></span>|<span data-ttu-id="7bdc3-185">管理者がアドインをアップロードして、ユーザーやグループに割り当てた。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-185">Admin uploaded the add-in and assigned it to users and/or groups.</span></span>|<span data-ttu-id="7bdc3-186">アドインが割り当てられたユーザーやグループは、関連する Office クライアントでアドインを表示できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-186">Users and/or groups assigned to the add-in see it in the relevant Office clients.</span></span>|
|<span data-ttu-id="7bdc3-187">**オフ**</span><span class="sxs-lookup"><span data-stu-id="7bdc3-187">**Turned off**</span></span>|<span data-ttu-id="7bdc3-188">管理者がアドインをオフにした。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-188">Admin turned off the add-in.</span></span>|<span data-ttu-id="7bdc3-p118">アドインに割り当てられたユーザーやグループは、アドインにアクセスできなくなります。 アドインの状態が **[オフ]** から **[アクティブ]** に変更されると、ユーザーやグループはアドインに再度アクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p118">Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.</span></span>|
|<span data-ttu-id="7bdc3-191">**Deleted**</span><span class="sxs-lookup"><span data-stu-id="7bdc3-191">**Deleted**</span></span>|<span data-ttu-id="7bdc3-192">管理者がアドインを削除した。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-192">Admin deleted the add-in.</span></span>|<span data-ttu-id="7bdc3-193">アドインが割り当てられたユーザーやグループは、そのアドインにアクセスできなくなります。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-193">Users and/or groups assigned the add-in no longer have access to it.</span></span>|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a><span data-ttu-id="7bdc3-194">一元展開によって発行された Office アドインの更新</span><span class="sxs-lookup"><span data-stu-id="7bdc3-194">Updating Office Add-ins that are published via Centralized Deployment</span></span>

<span data-ttu-id="7bdc3-p119">一元管理によってアドインが発行されると、アドインの Web アプリケーションに加えられた変更は、Web アプリケーションに変更が実装された直後に自動的にすべてのユーザーが使用できるようになります。 たとえば、アドインの [XML マニフェスト ファイル](../develop/add-in-manifests.md)に変更を加えると、アドインのアイコン、テキストまたはアドイン コマンドが次のように更新されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p119">After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:</span></span>

- <span data-ttu-id="7bdc3-p120">**基幹業務アドイン**: 管理者が Office 365 管理センターから一元展開を実施する際に明示的にマニフェストファイルをアップロードした場合は、管理者が目的の変更内容を含む新しいマニフェスト ファイルをアップロードする必要があります。 更新したマニフェスト ファイルがアップロードされると、関連する Office アプリケーションの次回起動時にアドインが更新されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p120">**Line-of-business add-in**: If an admin explicitly uploaded a manifest file when implementing Centralized Deployment via the Office 365 admin center, the admin must upload a new manifest file that contains the desired changes. After the updated manifest file has been uploaded, the next time the relevant Office applications start, the add-in will update.</span></span>

  > [!NOTE]
  > <span data-ttu-id="7bdc3-199">管理者は、更新を行うために LOB アドインを削除する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-199">An admin doesn't need to remove a LOB add-in to make an update.</span></span> <span data-ttu-id="7bdc3-200">[アドイン] セクションでは、管理者は LOB アドインを選択し、右下隅にある [**更新アドイン**] ボタンを押してこの機能を起動することができます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-200">In the Add-ins section, the admin can simply choose the LOB add-in and invoke this functionality by pressing the **Update add-in** button present in the bottom right corner.</span></span>
  > 
  > ![Office 365 管理センターの [アドインの更新] ダイアログを示すスクリーンショット](../images/update-add-in-admin-center.png)

- <span data-ttu-id="7bdc3-202">**Office ストア アドイン**: 管理者が Office 365 管理センターからの一元展開を実施するときに、Office ストアのアドインを選択した場合、アドインが Office ストアで更新されると、その後の一元展開によってアドインが更新されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-202">**Office Store add-in**: If an admin selected an add-in from the Office Store when implementing Centralized Deployment via the Office 365 admin center, and the add-in updates in the Office Store, the add-in will update later via Centralized Deployment.</span></span> <span data-ttu-id="7bdc3-203">関連する Office アプリケーションの次回起動時に、アドインが更新されます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-203">The next time the relevant Office applications start, the add-in will update.</span></span>

## <a name="end-user-experience-with-add-ins"></a><span data-ttu-id="7bdc3-204">アドインのエンド ユーザー エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="7bdc3-204">End user experience with add-ins</span></span>

<span data-ttu-id="7bdc3-205">一元展開によるアドインの発行が完了すると、エンド ユーザーはアドインがサポートする任意のプラットフォームでアドインの使用を開始できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-205">After an add-in has been published via Centralized Deployment, end users may start using it on any platform that the add-in supports.</span></span>

<span data-ttu-id="7bdc3-p123">アドインでアドイン コマンドがサポートされている場合、アドインが展開されているすべてのユーザーに対して、コマンドが Office アプリケーション リボンに表示されます。 次の例では、**[引用文献]** アドインのリボンに **[引用文献の検索]** コマンドが表示されています。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-p123">If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.</span></span>

![[引用文献] アドインの [引用文献の検索] コマンドが強調表示された Office リボンの部分を示すスクリーンショット](../images/search-citation.png)

<span data-ttu-id="7bdc3-209">アドインがアドイン コマンドをサポートしていない場合、ユーザーは次の手順を実行することで、Office アプリケーションにアドインを追加できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-209">If the add-in does not support add-in commands, users can add it to their Office application by doing the following:</span></span>

1. <span data-ttu-id="7bdc3-210">Word 2016 以降、Excel 2016 以降、または PowerPoint 2016 以降で **[挿入]** > **[個人用アドイン]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-210">In Word 2016 or later, Excel 2016 or later, or PowerPoint 2016 or later, choose **Insert** > **My Add-ins**.</span></span>
2. <span data-ttu-id="7bdc3-211">アドイン ウィンドウで **[管理者が管理]** タブを選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-211">Choose the **Admin Managed** tab in the add-in window.</span></span>
3. <span data-ttu-id="7bdc3-212">アドインを選択して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-212">Choose the add-in, and then choose **Add**.</span></span>

    ![Office アプリケーションの [Office アドイン] ページにある [管理者が管理] タブを示すスクリーンショット。 [引用文献] アドインがタブに表示されます。](../images/office-add-ins-admin-managed.png)

<span data-ttu-id="7bdc3-215">ただし、Outlook 2016 以降では、ユーザーは次の操作を実行できます。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-215">However, for Outlook 2016 or later, users can do the following:</span></span>

1. <span data-ttu-id="7bdc3-216">Outlook で **[ホーム]** > **[ストア]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-216">In Outlook, choose **Home** > **Store**.</span></span>
2. <span data-ttu-id="7bdc3-217">アドイン タブの下にある **[管理者が管理]** の項目を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-217">Choose the **Admin-managed** item under the add-in tab.</span></span>
3. <span data-ttu-id="7bdc3-218">アドインを選択して、**[追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bdc3-218">Choose the add-in, and then choose **Add**.</span></span>

    ![Outlook アプリケーションの [ストア] ページの [管理者が管理] 領域を示すスクリーンショット。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a><span data-ttu-id="7bdc3-220">関連項目</span><span class="sxs-lookup"><span data-stu-id="7bdc3-220">See also</span></span>

- [<span data-ttu-id="7bdc3-221">アドインの一元展開が Office 365 組織で動作するかどうかを判断する</span><span class="sxs-lookup"><span data-stu-id="7bdc3-221">Determine if Centralized Deployment of add-ins works for your Office 365 organization</span></span>](/office365/admin/manage/centralized-deployment-of-add-ins)
