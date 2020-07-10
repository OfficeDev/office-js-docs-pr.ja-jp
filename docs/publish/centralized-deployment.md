---
title: Microsoft 365 管理センターを介した一元展開を使用した Office アドインの発行
description: 一元展開を使用して、内部アドインと Isv が提供するアドインを展開する方法について説明します。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 0e99742be87b477b7c78295d08539de924f02466
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094254"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>Microsoft 365 管理センターを介した一元展開を使用した Office アドインの発行

Microsoft 365 管理センターを使用すると、管理者は組織内のユーザーやグループに Office アドインを簡単に展開できます。 管理センターを介して展開されたアドインは、Office アプリケーション内のユーザーがすぐに使用できるようになります。クライアント構成は必要ありません。 一元展開を使用して、Isv によって提供されるアドインと共に内部アドインを展開できます。

現在、Microsoft 365 管理センターは次のシナリオをサポートしています。

- 新しいアドインおよび更新されたアドインの個人、グループ、組織への一元展開。
- Windows、Mac、web を含む複数のクライアントプラットフォームへの展開。 Outlook では、iOS および Android への展開もサポートされています。 (ただし、iPad の Excel、Outlook、Word、PowerPoint のアドインのユーザーインストールはサポートされていますが、iPad への集中展開はサポートされて**いません**)。
- 英語および世界各国のテナントへの展開。
- クラウド ホスト型アドインの展開。
- ファイアウオール内でホストされているアドインの展開。
- AppSource アドインの展開。
- ユーザーのアドインの自動インストール (Office アプリケーション起動時)。
- ユーザーのアドインの自動削除 (管理者がアドインをオフにした場合や削除した場合。または、ユーザーが Azure Active Directory から削除された場合やアドインが展開されているグループから削除された場合)。

一元展開は、一元展開を使用するためのすべての要件を組織が満たしている場合に、Microsoft 365 管理者が組織内で Office アドインを展開するために推奨される方法です。 組織で集中型の展開を使用できるかどうかを判断する方法については、「 [Microsoft 365 組織でアドインの一元展開が機能](/office365/admin/manage/centralized-deployment-of-add-ins)するかどうかを判断する」を参照してください。

> [!NOTE]
> Microsoft 365 に接続されていないオンプレミス環境の場合、または Office 2013 を対象とした SharePoint アドインまたは Office アドインを展開する場合は、 [sharepoint アプリカタログ](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)を使用します。 COM/VSTO アドインを展開する場合は、ClickOnce または Windows インストーラーを使用してください。詳細については、「[Office ソリューションの配置](/visualstudio/vsto/deploying-an-office-solution)」を参照してください。

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Office アドインの展開に推奨されるアプローチ

Consider deploying Office Add-ins in a phased approach to help ensure that the deployment goes smoothly. We recommend the following plan:

1. Deploy the add-in to a small set of business stakeholders and members of the IT department. If the deployment is successful, move on to step 2.

2. Deploy the add-in to a larger set of individuals within the business who will be using the add-in. If the deployment is successful, move on to step 3.

3. アドインを使用することになるすべてのユーザーのグループにアドインを展開します。

対象ユーザーの規模に応じて、この手順に段階を追加するか、この手順から段階を削除してください。

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>一元展開による Office アドインの発行

開始する前に、「 [Microsoft 365 組織でのアドインの一元展開が機能するかどうかを判断](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)する」の説明に従って、組織が一元展開を使用するためのすべての要件を満たしていることを確認してください。

組織がすべての要件を満たしている場合は、次に示す手順を実行して、一元展開によって Office アドインを発行します。

1. 勤務先または教育機関のアカウントを使用して、Microsoft 365 にサインインします。
2. 左上にあるアプリ起動ツールのアイコンを選択して、**[管理]** をクリックします。
3. ナビゲーション メニューで、**[表示数を増やす]** を押し、**[設定]** > **[サービスとアドイン]** の順に選択します。
4. 新しい Microsoft 365 管理センターを発表するメッセージがページの上部に表示される場合は、メッセージを選択して管理センターのプレビューに移動します (「 [Microsoft 365 管理センターについ](/microsoft-365/admin/admin-overview/about-the-admin-center)て」を参照してください)。
5. ページの上部にある **[アドインの展開]** を選択します。
6. 要件の確認後、**[次へ]** を選択します。
7. **[一元展開]** ページで、次のいずれかのオプションを選択します。

    - **Office ストアからアドインを追加します。**
    - **I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![Microsoft 365 管理センターの [新しいアドイン] ダイアログ](../images/new-add-in.png)

8. Office ストアからアドインを追加するオプションを選択した場合は、アドインを選択します。 選択可能なアドインは、**[あなたへのおすすめ]**、**[評価]**、**[名前]** のカテゴリから表示できます。 Office ストアからは無料のアドインのみを追加できます。 有料のアドインの追加は、現在はサポートされていません。

    > [!NOTE]
    > Office ストアのオプションでは、管理者の操作なしで、ユーザーが自動的にアドインの更新と機能強化を利用できます。

    ![Microsoft 365 管理センターでアドインダイアログを選択する](../images/select-an-add-in.png)

9. アドインの詳細、プライバシーポリシー、ライセンス条項を確認した後、[**続行**] を選択します。

    ![Microsoft 365 管理センターで選択されているアドインページ](../images/selected-add-in-admin-center.png)

10. [**ユーザーの割り当て**] ページで、[**すべて**のユーザー]、[**特定のユーザー/グループ**]、または [**自分のみ**] を選択します。 検索ボックスを使用して、アドインを展開するユーザーやグループを検索します。 Outlook アドインの場合は、展開方法として**Fixed**、 **Available**、または**Optional**を選択することもできます。

    ![Microsoft 365 管理センターでアクセスと展開の方法を管理する](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > アドイン用の[シングル サインオン (SSO)](../develop/sso-in-office-add-ins.md) システムは現在プレビューなので、運用環境のアドインとして使用してはいけません。SSO を使用するアドインが展開されている場合、割り当てられているユーザーとグループは、同じ Azure アプリ ID を共有するアドインによっても共有されます。 ユーザーの割り当ての変更は、これらのアドインにも適用されます。関連するアドインは、このページに表示されます。 SSO アドインに限り、アドインで必要な Microsoft Graph アクセス許可のリストがこのページに表示されます。

11. 完了したら、[**展開**] を選択します。 このプロセスには、最大で 3 分かかる場合があります。 その後、**[次へ]** を押してチュートリアルを終了します。 アドインが Office 365 のその他のアプリと共に表示されるようになります。

    > [!NOTE]
    > 管理者が [**展開**] を選択すると、すべてのユーザーに同意が与えられます。

    ![Microsoft 365 管理センターのアプリの一覧](../images/citations.png)

> [!TIP]
> 組織内のユーザーやグループに新しいアドインを展開するときには、いつどのようにアドインを使用するかについての説明と、サポート資料 (関連するヘルプ コンテンツやよくある質問など) へのリンクを含む電子メールの送信を検討してください。

## <a name="considerations-when-granting-access-to-an-add-in"></a>アドインへのアクセスを許可するときの考慮事項

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:

- **Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

- **Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.

- **グループ**: グループにアドインを割り当てると、グループに追加されたユーザーにアドインが自動的に割り当てられます。 同様に、ユーザーがグループから削除されると、そのユーザーはアドインへのアクセス権を自動的に失います。 どちらの場合も、Microsoft 365 admin から追加の操作を行う必要はありません。

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be more practical to assign the add-in to specific users.

## <a name="add-in-states"></a>アドインの状態

次の表では、様々なアドインの状態について説明しています。

|状態|状態が発生する原因|影響|
|-----|--------------------|------|
|**アクティブ**|管理者がアドインをアップロードして、ユーザーやグループに割り当てた。|アドインが割り当てられたユーザーやグループは、関連する Office クライアントでアドインを表示できます。|
|**オフ**|管理者がアドインをオフにした。|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will regain access to it.|
|**Deleted**|管理者がアドインを削除した。|アドインが割り当てられたユーザーやグループは、そのアドインにアクセスできなくなります。|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>一元展開によって発行された Office アドインの更新

After an Office Add-in has been published via Centralized Deployment, any changes made to the add-in's web application will automatically be available to all users as soon as those changes are implemented in the web application. Changes made to an add-in's [XML manifest file](../develop/add-in-manifests.md), for example, to update the add-in's icon, text, or add-in commands, happen as follows:

- **基幹業務アドイン**: 管理者が、Microsoft 365 管理センターを使用した集中型展開の実装時にマニフェストファイルを明示的にアップロードした場合、管理者は必要な変更を含む新しいマニフェストファイルをアップロードする必要があります。 更新したマニフェスト ファイルがアップロードされると、関連する Office アプリケーションの次回起動時にアドインが更新されます。

  > [!NOTE]
  > 管理者は、更新を行うために LOB アドインを削除する必要はありません。 [アドイン] セクションでは、管理者は LOB アドインを選択し、右下隅にある [**更新アドイン**] ボタンを押してこの機能を起動することができます。
  > 
  > ![Microsoft 365 管理センターの [アドインの更新] ダイアログを示すスクリーンショット](../images/update-add-in-admin-center.png)

- **Office ストアアドイン**: 管理者が Microsoft 365 管理センターを使用して一元展開を実装するときに office ストアからアドインを選択し、office ストアでアドインを更新すると、そのアドインは後で一元展開によって更新されます。 関連する Office アプリケーションの次回起動時に、アドインが更新されます。

## <a name="end-user-experience-with-add-ins"></a>アドインのエンド ユーザー エクスペリエンス

一元展開によるアドインの発行が完了すると、エンド ユーザーはアドインがサポートする任意のプラットフォームでアドインの使用を開始できます。

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. In the following example, the command **Search Citation** appears in the ribbon for the **Citations** add-in.

![[引用] アドインの [引用文献の検索] コマンドが強調表示されている Office アプリリボンのセクションを示すスクリーンショット](../images/search-citation.png)

アドインがアドイン コマンドをサポートしていない場合、ユーザーは次の手順を実行することで、Office アプリケーションにアドインを追加できます。

1. Word 2016 以降、Excel 2016 以降、または PowerPoint 2016 以降で **[挿入]** > **[個人用アドイン]** の順に選択します。
2. アドイン ウィンドウで **[管理者が管理]** タブを選択します。
3. アドインを選択して、**[追加]** を選択します。

    ![Screenshot shows the Admin Managed tab of the Office Add-ins page of an Office application. The Citations add-in is shown on the tab.](../images/office-add-ins-admin-managed.png)

ただし、Outlook 2016 以降では、ユーザーは次の操作を実行できます。

1. Outlook で **[ホーム]** > **[ストア]** の順に選択します。
2. アドイン タブの下にある **[管理者が管理]** の項目を選択します。
3. アドインを選択して、**[追加]** を選択します。

    ![Outlook アプリケーションの [ストア] ページの [管理者が管理] 領域を示すスクリーンショット。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>関連項目

- [アドインの一元展開が Microsoft 365 組織で動作するかどうかを判断する](/office365/admin/manage/centralized-deployment-of-add-ins)
