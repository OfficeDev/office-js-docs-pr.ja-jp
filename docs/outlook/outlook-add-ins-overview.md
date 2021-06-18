---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 06/15/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: f0c1dbdd1cf9909310b629188d4f3d3d5de6b6bb
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007812"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="00df7-103">Outlook アドインの概要</span><span class="sxs-lookup"><span data-stu-id="00df7-103">Outlook add-ins overview</span></span>

<span data-ttu-id="00df7-p101">Outlook アドインは、Web ベースのプラットフォームを使用してサードパーティ企業によって Outlook に組み込まれた統合機能です。Outlook アドインには次の 3 つの主な側面があります。</span><span class="sxs-lookup"><span data-stu-id="00df7-p101">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="00df7-106">Windows と Mac 用のデスクトップ Outlook、Web 版 (Microsoft 365 と Outlook.com)、モバイル版すべてで機能する同じアドインとビジネス ロジック。</span><span class="sxs-lookup"><span data-stu-id="00df7-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="00df7-107">Outlook アドインは、マニフェスト (アドインが Outlook に統合する方法 (ボタンや作業ウィンドウなど) を説明する)、および JavaScript/HTML のコード (アドインの UI とビジネス ロジックを構成する) で構成される。</span><span class="sxs-lookup"><span data-stu-id="00df7-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="00df7-108">Outlook アドインは、[AppSource](https://appsource.microsoft.com) から入手するか、エンドユーザーまたは管理者が[サイドロード](sideload-outlook-add-ins-for-testing.md)することができます。</span><span class="sxs-lookup"><span data-stu-id="00df7-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="00df7-p102">Outlook アドインは、Windows で実行する Outlook に固有の古い統合である COM アドインや VSTO アドインとは異なります。COM アドインとは異なり、Outlook アドインには、ユーザーのデバイスや Outlook クライアントに物理的にインストールされたコードがありません。Outlook アドインの場合、Outlook はマニフェストを読み取り、指定された UI コントロールをフックして、JavaScript と HTML を読み込みます。Web コンポーネントは全て、サンドボックス内のブラウザーのコンテキストで実行されます。</span><span class="sxs-lookup"><span data-stu-id="00df7-p102">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="00df7-p103">アドインをサポートする Outlook アイテムには、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消し、予定などがあります。それぞれの Outlook アドインでは、メール アドインが使用できるコンテキストを定義します。これにはアイテムの種類、およびユーザーがアイテムの読み取り (または作成) を行っているかどうかなどがあります。</span><span class="sxs-lookup"><span data-stu-id="00df7-p103">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="00df7-115">拡張点</span><span class="sxs-lookup"><span data-stu-id="00df7-115">Extension points</span></span>

<span data-ttu-id="00df7-p104">拡張点は、アドインが Outlook と統合する方法です。これを行う方法は以下のとおりです。</span><span class="sxs-lookup"><span data-stu-id="00df7-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="00df7-p105">アドインは、メッセージと予定のコマンド サーフェスに表示されるボタンを宣言できます。詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="00df7-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="00df7-120">**リボン上の [コマンド] ボタンがあるアドイン**</span><span class="sxs-lookup"><span data-stu-id="00df7-120">**An add-in with command buttons on the ribbon**</span></span>

    ![アドイン コマンドの UI なし図形](../images/uiless-command-shape.png)

- <span data-ttu-id="00df7-p106">アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="00df7-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="00df7-124">**強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**</span><span class="sxs-lookup"><span data-stu-id="00df7-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![カード内のコンテキスト アプリを示す](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="00df7-126">アドインで使用可能なメールボックスのアイテム</span><span class="sxs-lookup"><span data-stu-id="00df7-126">Mailbox items available to add-ins</span></span>

<span data-ttu-id="00df7-127">Outlook アドインは、ユーザーがメッセージまたは予定を作成または読んでいるときにアクティブになりますが、他の種類のアイテムではアクティブになりません。</span><span class="sxs-lookup"><span data-stu-id="00df7-127">Outlook add-ins activate when the user is composing or reading a message or appointment, but not other item types.</span></span> <span data-ttu-id="00df7-128">ただし、現在のメッセージ アイテムが作成または読み取りフォームで次のいずれかである場合、アドインはアクティブ化 *されません*。</span><span class="sxs-lookup"><span data-stu-id="00df7-128">However, add-ins are *not* activated if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="00df7-p108">Information Rights Management (IRM) によって保護されているか、または保護のためにその他の方法で暗号化されている場合。デジタル署名はこれらいずれかのメカニズムに依存しているため、デジタル署名されたメッセージはその一例です。</span><span class="sxs-lookup"><span data-stu-id="00df7-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  >
  > - <span data-ttu-id="00df7-131">アドインは、Microsoft 365 サブスクリプションに関連付けられている Outlook のデジタル署名付きメッセージでライセンス認証を行います。</span><span class="sxs-lookup"><span data-stu-id="00df7-131">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="00df7-132">Windows では、このサポートはビルド 8711.1000 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="00df7-132">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="00df7-133">Windows の Outlook ビルド 13229.10000 から、IRM で保護されたアイテムに対してアドインをアクティブ化できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="00df7-133">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="00df7-134">この機能のプレビューの詳細については、「[Information Rights Management (IRM) で保護されているアイテムのアドインのアクティブ化](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="00df7-134">For more information about this feature in preview, refer to [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="00df7-135">メッセージ クラスが IPM.Report.\* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。</span><span class="sxs-lookup"><span data-stu-id="00df7-135">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="00df7-136">別のメッセージに添付される .msg または .eml ファイルの場合。</span><span class="sxs-lookup"><span data-stu-id="00df7-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="00df7-137">.msg または .eml ファイルがファイル システムから開かれた場合。</span><span class="sxs-lookup"><span data-stu-id="00df7-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="00df7-138">共有メールボックス\*の[グループ メールボックス](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)内、別のユーザーのメールボックス内\*、アーカイブ メールボックス内、パブリック フォルダー内。</span><span class="sxs-lookup"><span data-stu-id="00df7-138">In a [group mailbox](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), in a shared mailbox\*, in another user's mailbox\*, in an archive mailbox, or in a public folder.</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="00df7-139">\* [要件セット 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) では、代理アクセス シナリオ (別のユーザーのメールボックスで共有されるフォルダなど) のサポートが導入されました。</span><span class="sxs-lookup"><span data-stu-id="00df7-139">\* Support for delegate access scenarios (for example, folders shared from another user's mailbox) was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="00df7-140">共有メールボックスのサポートをプレビューしています。</span><span class="sxs-lookup"><span data-stu-id="00df7-140">Shared mailbox support is now in preview.</span></span> <span data-ttu-id="00df7-141">詳細については、「[共有フォルダーと共有メールボックスのシナリオを有効にする](delegate-access.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="00df7-141">To learn more, refer to [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

- <span data-ttu-id="00df7-142">カスタム フォームを使用する場合。</span><span class="sxs-lookup"><span data-stu-id="00df7-142">Using a custom form.</span></span>

<span data-ttu-id="00df7-143">既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="00df7-143">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="00df7-144">この理由の詳細は、[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)の「既知のエンティティに対するサポート」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="00df7-144">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-clients"></a><span data-ttu-id="00df7-145">サポートされるクライアント</span><span class="sxs-lookup"><span data-stu-id="00df7-145">Supported clients</span></span>

<span data-ttu-id="00df7-146">Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Outlook on the web と Outlook.com でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="00df7-146">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="00df7-147">最新の機能すべてが、すべての[クライアント](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="00df7-147">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="00df7-148">これらの機能が各アプリケーションでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="00df7-148">Please refer to articles and API references for those features to see which applications they may or may not be supported in.</span></span>

## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="00df7-149">Outlook アドインの作成を開始する</span><span class="sxs-lookup"><span data-stu-id="00df7-149">Get started building Outlook add-ins</span></span>

<span data-ttu-id="00df7-150">Outlook アドインの作成を開始するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="00df7-150">To get started building Outlook add-ins, try the following:</span></span>

- <span data-ttu-id="00df7-151">[クイックスタート](../quickstarts/outlook-quickstart.md) - 簡単な作業ウィンドウを作成します。</span><span class="sxs-lookup"><span data-stu-id="00df7-151">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="00df7-152">[チュートリアル](../tutorials/outlook-tutorial.md) - 新しいメッセージに GitHub gist を挿入するアドインを作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="00df7-152">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>

## <a name="see-also"></a><span data-ttu-id="00df7-153">関連項目</span><span class="sxs-lookup"><span data-stu-id="00df7-153">See also</span></span>

- [<span data-ttu-id="00df7-154">Microsoft 365 開発者プログラムについて</span><span class="sxs-lookup"><span data-stu-id="00df7-154">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="00df7-155">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="00df7-155">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="00df7-156">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="00df7-156">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="00df7-157">Office および SharePoint アドインのライセンスを付与する</span><span class="sxs-lookup"><span data-stu-id="00df7-157">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="00df7-158">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="00df7-158">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="00df7-159">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="00df7-159">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
