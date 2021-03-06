---
title: Office 365 管理センターからの一元展開を使用した Office アドインの発行
description: 集中展開を使用して、ISV によって提供される内部アドインとアドインを展開する方法について説明します。
ms.date: 03/22/2021
localization_priority: Normal
ms.openlocfilehash: b57e21f177fe66f03985ce6baee4d9eeda75d8bd
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348791"
---
# <a name="publish-office-add-ins-using-centralized-deployment-via-the-microsoft-365-admin-center"></a>Office 365 管理センターからの一元展開を使用した Office アドインの発行

このMicrosoft 365 管理センターを使用すると、管理者は組織内のユーザー Officeにアドインを簡単に展開できます。 管理センター経由で展開されたアドインは、ユーザーがすぐに Office アプリケーションで利用できるようになります。クライアントの構成は必要ありません。 一元展開は、内部アドインの展開に使用することも、ISV が提供するアドインの展開に使用することもできます。

現在、Microsoft 365 管理センターは次のシナリオをサポートしています。

- 新しいアドインおよび更新されたアドインの個人、グループ、組織への一元展開。
- 複数のクライアント プラットフォームへの展開 (Windows Mac、Web など)。 このOutlook、iOS と Android への展開もサポートされています。 (ただし、iPad での Excel、Outlook、Word、および PowerPoint アドインのユーザー インストールはサポートされていますが、iPad への一元展開 **はサポートされていません**。
- 英語および世界各国のテナントへの展開。
- クラウド ホスト型アドインの展開。
- ファイアウオール内でホストされているアドインの展開。
- AppSource アドインの展開。
- ユーザーのアドインの自動インストール (Office アプリケーション起動時)。
- ユーザーのアドインの自動削除 (管理者がアドインをオフにした場合や削除した場合。または、ユーザーが Azure Active Directory から削除された場合やアドインが展開されているグループから削除された場合)。

組織が集中展開を使用するためのすべての要件を満たしている場合、Microsoft 365 管理者が組織内に Office アドインを展開するための推奨される方法は、集中展開です。 組織で集中展開を使用できるかどうかを判断する方法については、「アドインの集中展開が組織に対して機能するかどうかを判断する」[をMicrosoft 365してください](/office365/admin/manage/centralized-deployment-of-add-ins)。

> [!NOTE]
> Microsoft 365 への接続がないオンプレミス環境、または Office 2013 を対象とする SharePoint アドインまたは Office アドインを展開するには、SharePoint アプリ カタログを使用[します](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)。 COM/VSTO アドインを展開する場合は、ClickOnce または Windows インストーラーを使用してください。詳細については、「[Office ソリューションの配置](/visualstudio/vsto/deploying-an-office-solution)」を参照してください。

## <a name="recommended-approach-for-deploying-office-add-ins"></a>Office アドインの展開に推奨されるアプローチ

展開がスムーズに進行するように、段階的なアプローチで Office アドインを展開することを検討してください。 以下のプランをお勧めします。

1. ビジネス関係者の少人数のグループと IT 部門のメンバーにアドインを展開します。 展開が成功した場合は、第 2 段階に進みます。

1. アドインを使用することになるビジネス ユーザーの人数を増やしたグループにアドインを展開します。 展開が成功した場合は、第 3 段階に進みます。

1. アドインを使用することになるすべてのユーザーのグループにアドインを展開します。

対象ユーザーの規模に応じて、この手順に段階を追加するか、この手順から段階を削除してください。

## <a name="publish-an-office-add-in-via-centralized-deployment"></a>一元展開による Office アドインの発行

開始する前に、「アドインの集中展開が Microsoft 365 組織で機能するかどうかを判断する」の説明に従って、組織が集中展開の使用に関するすべての要件を満たしていることを[確認](/microsoft-365/admin/manage/centralized-deployment-of-add-ins)します。

組織がすべての要件を満たしている場合は、次の手順を実行して、一元展開をOfficeアドインを発行します。

1. 仕事または教育アカウントMicrosoft 365にサインインします。
1. 左上にあるアプリ起動ツールのアイコンを選択して、**[管理]** をクリックします。
1. ナビゲーション メニューで、[詳細の表示]**を** 選択し、[統合  >  **設定] を選択します**。
1. ページの上部で、[アドイン] **を選択します**。
1. 新しい Microsoft 365 管理センター を発表するメッセージがページの上部に表示される場合は、管理センター プレビューに移動するメッセージを選択します (「概要」を[参照](/microsoft-365/admin/admin-overview/about-the-admin-center)Microsoft 365 管理センター)。
1. ページの上部にある **[アドインの展開]** を選択します。
1. 要件の確認後、**[次へ]** を選択します。
1. [集中展開] ページで、次のいずれかの **オプションを選択** します。

    - **Office ストアからアドインを追加します。**
    - **このデバイスにマニフェスト ファイル (.xml) があります。** このオプションの場合は、**[参照]** を選択して、使用するマニフェスト ファイル (.xml) の場所を指定します。
    - **マニフェスト ファイルの URL がわかります。** このオプションの場合は、所定のフィールドにマニフェストの URL を入力します。

    ![[Add-In] ダイアログ ボックスMicrosoft 365 管理センター。](../images/new-add-in.png)

1. Office ストアからアドインを追加するオプションを選択した場合は、アドインを選択します。 選択可能なアドインは、**[あなたへのおすすめ]**、**[評価]**、**[名前]** のカテゴリから表示できます。 Office ストアからは無料のアドインのみを追加できます。 有料のアドインの追加は、現在はサポートされていません。

    > [!NOTE]
    > Office ストアのオプションでは、管理者の操作なしで、ユーザーが自動的にアドインの更新と機能強化を利用できます。

    ![[アドイン] ダイアログボックスを選択Microsoft 365 管理センター。](../images/select-an-add-in.png)

1. アドイン **の詳細** 、プライバシー ポリシー、ライセンス条項を確認した後、[続行] を選択します。

    ![[アドイン] ページで選択Microsoft 365 管理センター。](../images/selected-add-in-admin-center.png)

1. [ユーザーの **割り当て] ページ** で、[ **すべてのユーザー**] 、[ **特定のユーザー/グループ]**、または [自分のみ] **を選択します**。 検索ボックスを使用して、アドインを展開するユーザーやグループを検索します。 アドインOutlook、展開方法 [固定] 、[利用可能]、または **[オプション**] を選択 **できます**。

    ![アクセスおよび展開方法を持つユーザーを管理Microsoft 365 管理センター。](../images/manage-users-deployment-admin-center.png)

    > [!NOTE]
    > シングル サインオン [(SSO)](../develop/sso-in-office-add-ins.md) を使用するアドインは、アドイン マニフェストに記載されているスコープに同意するよう管理者に求めるメッセージを表示します。  同じバッキング サービスが複数のアドインで使用されている場合 (同じ Azure App ID が異なるアドインの SSO で使用される)、各アドインのスコープに対して、各展開に対する同意を求めるメッセージが表示されます。 このページには、アドインで必要なアクセス許可の一覧も表示されます。

1. 完了したら、[展開] **を選択します**。 このプロセスには、最大で 3 分かかる場合があります。 その後、**[次へ]** を押してチュートリアルを終了します。 これで、他のアプリと共にアドインがOfficeされます。

    > [!NOTE]
    > 管理者が [展開] を **選択すると**、すべてのユーザーに同意が与えられる。

    ![アプリの一覧 (Microsoft 365 管理センター)。](../images/citations.png)

> [!TIP]
> 組織内のユーザーやグループに新しいアドインを展開するときには、いつどのようにアドインを使用するかについての説明と、サポート資料 (関連するヘルプ コンテンツやよくある質問など) へのリンクを含む電子メールの送信を検討してください。

## <a name="considerations-when-granting-access-to-an-add-in"></a>アドインへのアクセスを許可するときの考慮事項

管理者は、組織内のすべてのユーザーにアドインを割り当てることも、組織内の特定のユーザーやグループにアドインを割り当てることもできます。 次の一覧では、各オプションの影響について説明します。

- **すべてのユーザー**: 名前が示すように、このオプションでは、テナント内のすべてのユーザーにアドインが割り当てられます。対象のアドインが組織全体で汎用な場合にのみ、このオプションを慎重に使用します。

- **ユーザー**: 個別のユーザーにアドインを割り当てる場合は、追加のユーザーにアドインを割り当てるたびに、アドインの一元展開の設定を更新する必要があります。 同様に、ユーザーのアドインへのアクセス権を削除するたびに、アドインの一元展開の設定を更新する必要があります。

- **グループ**: グループにアドインを割り当てると、グループに追加されたユーザーにアドインが自動的に割り当てられます。 同様に、ユーザーがグループから削除されると、そのユーザーはアドインへのアクセス権を自動的に失います。 いずれの場合も、管理者から追加のアクションはMicrosoft 365ありません。

一般に、保守が簡単になるように、可能な場合は常にグループを使用してアドインを割り当てるようにしてください。 ただし、アドインのアクセスをユーザーの非常に少数のメンバーに制限する場合は、具体的なユーザーにアドインを割り当てるほうが実用的です。

## <a name="add-in-states"></a>アドインの状態

次の表では、様々なアドインの状態について説明しています。

|状態|状態が発生する原因|影響|
|-----|--------------------|------|
|**アクティブ**|管理者がアドインをアップロードして、ユーザーやグループに割り当てた。|アドインが割り当てられたユーザーやグループは、関連する Office クライアントでアドインを表示できます。|
|**オフ**|管理者がアドインをオフにした。|アドインに割り当てられたユーザーやグループは、アドインにアクセスできなくなります。 アドインの状態が **[オフ]** から **[アクティブ]** に変更されると、ユーザーやグループはアドインに再度アクセスできるようになります。|
|**Deleted**|管理者がアドインを削除した。|アドインが割り当てられたユーザーやグループは、そのアドインにアクセスできなくなります。|

## <a name="updating-office-add-ins-that-are-published-via-centralized-deployment"></a>一元展開によって発行された Office アドインの更新

Officeを一元的な展開で公開すると、その変更が Web アプリケーションに実装された後、アドインの Web アプリケーションに加えた変更は自動的にすべてのユーザーが利用できます。 たとえば、アドインのアイコン、テキスト、アドイン のコマンドを更新するためにアドインの [XML](../develop/add-in-manifests.md) マニフェスト ファイルに加えた変更は、次のように行われます。

- **Line-of-business** アドイン : Microsoft 365 管理センター 経由で集中展開を実装する際に、管理者がマニフェスト ファイルを明示的にアップロードした場合 (デバイスまたは URL を指す) 場合、管理者は目的の変更を含む新しいマニフェスト ファイルをアップロードする必要があります。 更新したマニフェスト ファイルがアップロードされると、関連する Office アプリケーションの次回起動時にアドインが更新されます。

  > [!NOTE]
  > 管理者は、更新を行う LOB アドインを削除する必要があります。 [アドイン] セクションでは、管理者は LOB アドインを選択するだけで、右下隅にある [アドインの更新]ボタンを押してこの機能を呼び出します。
  >
  > ![スクリーンショットは、[アドインの更新] ダイアログを表示Microsoft 365 管理センター。](../images/update-add-in-admin-center.png)

- **Office** ストア アドイン : Microsoft 365 管理センター 経由で集中展開を実装するときに管理者が Office ストアからアドインを選択し、Office ストアでアドインの更新を行った場合、アドインは後で集中展開を介して更新されます。 すべてのエンド ユーザーに対してストア アドインの更新プログラムがフローするには、最大 24 時間かかる場合があります。 この期間が経過すると、Officeアプリケーションが再起動すると、アドインが更新されます。 ユーザーは、[タブ アドインの挿入] [管理] タブの [ヒット更新] を選択して、手動更新をトリガーして最新のストア アドイン バージョン  >    >    >  **を取得することもできます**。

## <a name="end-user-experience-with-add-ins"></a>アドインのエンド ユーザー エクスペリエンス

一元展開によるアドインの発行が完了すると、エンド ユーザーはアドインがサポートする任意のプラットフォームでアドインの使用を開始できます。

アドインでアドイン コマンドがサポートされている場合、アドインが展開されているすべてのユーザーに対して、コマンドが Office アプリケーション リボンに表示されます。 次の例では、**[引用文献]** アドインのリボンに **[引用文献の検索]** コマンドが表示されています。

![スクリーンショットは、[引用文献] アドインOffice アプリ[引用文献の検索] コマンドが強調表示されたリボンのセクションを示しています。](../images/search-citation.png)

アドインがアドイン コマンドをサポートしていない場合、ユーザーは次の手順を実行することで、Office アプリケーションにアドインを追加できます。

1. Word 2016 以降、Excel 2016 以降、または PowerPoint 2016 以降で **[挿入]** > **[個人用アドイン]** の順に選択します。
1. アドイン ウィンドウで **[管理者が管理]** タブを選択します。
1. アドインを選択して、**[追加]** を選択します。

    ![Office アプリケーションの [Office アドイン] ページにある [管理者が管理] タブを示すスクリーンショット。 [引用文献] アドインがタブに表示されます。](../images/office-add-ins-admin-managed.png)

ただし、Outlook 2016 以降では、ユーザーは次の操作を実行できます。

1. Outlook で **[ホーム]** > **[ストア]** の順に選択します。
1. アドイン タブの下にある **[管理者が管理]** の項目を選択します。
1. アドインを選択して、**[追加]** を選択します。

    ![Outlook アプリケーションの [ストア] ページの [管理者が管理] 領域を示すスクリーンショット。](../images/outlook-add-ins-admin-managed.png)

## <a name="see-also"></a>関連項目

- [アドインの一元展開が Microsoft 365 組織で動作するかどうかを判断する](/office365/admin/manage/centralized-deployment-of-add-ins)
