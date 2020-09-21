---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 09/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 09f410ccbddb4cffadc700036a4da3c45d2fb6e3
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819568"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="4198d-103">Outlook アドインの概要</span><span class="sxs-lookup"><span data-stu-id="4198d-103">Outlook add-ins overview</span></span>

<span data-ttu-id="4198d-104">Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。</span><span class="sxs-lookup"><span data-stu-id="4198d-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="4198d-105">Outlook アドインには次の 3 つの主な側面があります。</span><span class="sxs-lookup"><span data-stu-id="4198d-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="4198d-106">Windows と Mac 用のデスクトップ Outlook、Web 版 (Microsoft 365 と Outlook.com)、モバイル版すべてで機能する同じアドインとビジネス ロジック。</span><span class="sxs-lookup"><span data-stu-id="4198d-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="4198d-107">Outlook アドインは、マニフェスト (アドインが Outlook に統合する方法 (ボタンや作業ウィンドウなど) を説明する)、および JavaScript/HTML のコード (アドインの UI とビジネス ロジックを構成する) で構成される。</span><span class="sxs-lookup"><span data-stu-id="4198d-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="4198d-108">Outlook アドインは、[AppSource](https://appsource.microsoft.com) から入手するか、エンドユーザーまたは管理者が[サイドロード](sideload-outlook-add-ins-for-testing.md)することができます。</span><span class="sxs-lookup"><span data-stu-id="4198d-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="4198d-109">Outlook アドインは、Windows 版 Outlook 固有の統合機能として以前から存在した COM アドインや VSTO アドインとは異なります。</span><span class="sxs-lookup"><span data-stu-id="4198d-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="4198d-110">COM アドインとは違い、Outlook アドインのコードがユーザーのデバイスまたは Outlook クライアントに物理的にインストールされることはありません。</span><span class="sxs-lookup"><span data-stu-id="4198d-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="4198d-111">Outlook のアドインの場合、Outlook はマニフェストを読み取り UI で指定したコントロールをフックした後に、HTML と JavaScript を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="4198d-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="4198d-112">この Web コンポーネントは、サンドボックス内のブラウザーのコンテキストですべて実行されます。</span><span class="sxs-lookup"><span data-stu-id="4198d-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="4198d-113">アドインをサポートしている Outlook アイテムには、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消し、予定などがあります。</span><span class="sxs-lookup"><span data-stu-id="4198d-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="4198d-114">それぞれの Outlook アドインにより、アイテムの種類、ユーザーがアイテムの読み取りや作成を行うかどうかなど、使用できるコンテキストが定義されます。</span><span class="sxs-lookup"><span data-stu-id="4198d-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="4198d-115">拡張点</span><span class="sxs-lookup"><span data-stu-id="4198d-115">Extension points</span></span>

<span data-ttu-id="4198d-p104">拡張点は、アドインが Outlook と統合する方法です。これを行う方法は以下のとおりです。</span><span class="sxs-lookup"><span data-stu-id="4198d-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="4198d-p105">アドインは、メッセージと予定のコマンド サーフェスに表示されるボタンを宣言できます。詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4198d-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="4198d-120">**リボン上の [コマンド] ボタンがあるアドイン**</span><span class="sxs-lookup"><span data-stu-id="4198d-120">**An add-in with command buttons on the ribbon**</span></span>

    ![アドイン コマンドの UI なし図形](../images/uiless-command-shape.png)

- <span data-ttu-id="4198d-p106">アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4198d-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="4198d-124">**強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**</span><span class="sxs-lookup"><span data-stu-id="4198d-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![カード内のコンテキスト アプリを示す](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="4198d-126">アドインで使用可能なメールボックス アイテム</span><span class="sxs-lookup"><span data-stu-id="4198d-126">Mailbox items available to add-ins</span></span>

<span data-ttu-id="4198d-p107">Outlook アドインは、作成中や読み取り中にメッセージや予定で使用することができますが、他のアイテムの種類では使用できません。新規作成フォームまたは閲覧フォームで現在のメッセージ アイテムが次のいずれかの場合、Outlook はアドインをアクティブ化しません。</span><span class="sxs-lookup"><span data-stu-id="4198d-p107">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="4198d-p108">Information Rights Management (IRM) によって保護されているか、または保護のためにその他の方法で暗号化されている場合。デジタル署名はこれらいずれかのメカニズムに依存しているため、デジタル署名されたメッセージはその一例です。</span><span class="sxs-lookup"><span data-stu-id="4198d-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  > - <span data-ttu-id="4198d-131">アドインは、Microsoft 365 サブスクリプションに関連付けられている Outlook のデジタル署名付きメッセージでライセンス認証を行います。</span><span class="sxs-lookup"><span data-stu-id="4198d-131">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="4198d-132">Windows では、このサポートはビルド 8711.1000 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="4198d-132">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="4198d-133">Windows の Outlook ビルド 13229.10000 から、IRM で保護されたアイテムに対してアドインをアクティブ化できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="4198d-133">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="4198d-134">この機能のプレビューの詳細については、「[Information Rights Management (IRM) で保護されているアイテムのアドインのアクティブ化](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4198d-134">For more information about this feature in preview, see [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="4198d-135">メッセージ クラスが IPM.Report.\* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。</span><span class="sxs-lookup"><span data-stu-id="4198d-135">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="4198d-136">下書きであるか (送信者が割り当てられていない)、Outlook の [下書き] フォルダーにある場合。</span><span class="sxs-lookup"><span data-stu-id="4198d-136">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="4198d-137">別のメッセージに添付される .msg または .eml ファイルの場合。</span><span class="sxs-lookup"><span data-stu-id="4198d-137">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="4198d-138">.msg または .eml ファイルがファイル システムから開かれた場合。</span><span class="sxs-lookup"><span data-stu-id="4198d-138">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="4198d-139">共有メールボックス内、別のユーザーのメールボックス内、アーカイブ メールボックス内、パブリック フォルダー内。</span><span class="sxs-lookup"><span data-stu-id="4198d-139">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="4198d-140">カスタム フォームを使用する場合。</span><span class="sxs-lookup"><span data-stu-id="4198d-140">Using a custom form.</span></span>

<span data-ttu-id="4198d-141">既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="4198d-141">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="4198d-142">この理由の詳細は、[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)の「既知のエンティティに対するサポート」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="4198d-142">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-clients"></a><span data-ttu-id="4198d-143">サポートされるクライアント</span><span class="sxs-lookup"><span data-stu-id="4198d-143">Supported clients</span></span>

<span data-ttu-id="4198d-144">Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Outlook on the web と Outlook.com でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="4198d-144">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="4198d-145">最新の機能すべてが、すべての[クライアント](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="4198d-145">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="4198d-146">これらの機能が各アプリケーションでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="4198d-146">Please refer to articles and API references for those features to see which applications they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="4198d-147">Outlook アドインの作成を開始する</span><span class="sxs-lookup"><span data-stu-id="4198d-147">Get started building Outlook add-ins</span></span>

<span data-ttu-id="4198d-148">Outlook アドインの作成を開始するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="4198d-148">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="4198d-149">[クイックスタート](../quickstarts/outlook-quickstart.md) - 簡単な作業ウィンドウを作成します。</span><span class="sxs-lookup"><span data-stu-id="4198d-149">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="4198d-150">[チュートリアル](../tutorials/outlook-tutorial.md) - 新しいメッセージに GitHub gist を挿入するアドインを作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="4198d-150">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="4198d-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="4198d-151">See also</span></span>

- [<span data-ttu-id="4198d-152">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="4198d-152">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="4198d-153">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="4198d-153">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="4198d-154">Office および SharePoint アドインのライセンスを付与する</span><span class="sxs-lookup"><span data-stu-id="4198d-154">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="4198d-155">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="4198d-155">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="4198d-156">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="4198d-156">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
