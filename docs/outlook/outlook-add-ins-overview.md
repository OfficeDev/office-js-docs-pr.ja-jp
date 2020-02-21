---
title: Outlook アドインの概要
description: Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。
ms.date: 10/09/2019
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: cb6e19788390a804b0bbacb97666a3ca8a9d5971
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166561"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="1d9f9-103">Outlook アドインの概要</span><span class="sxs-lookup"><span data-stu-id="1d9f9-103">Outlook add-ins overview</span></span>

<span data-ttu-id="1d9f9-104">Outlook アドインとは、Microsoft の Web ベース プラットフォームを使用して Outlook に組み込まれるサードパーティ製の統合機能です。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="1d9f9-105">Outlook アドインには次の 3 つの主な側面があります。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="1d9f9-106">Windows と Mac 用のデスクトップ Outlook、Web 版 (Office 365 と Outlook.com)、モバイル版すべてで機能する同じアドインとビジネス ロジック。 </span><span class="sxs-lookup"><span data-stu-id="1d9f9-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Office 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="1d9f9-107">Outlook アドインは、マニフェスト (アドインが Outlook に統合する方法 (ボタンや作業ウィンドウなど) を説明する)、および JavaScript/HTML のコード (アドインの UI とビジネス ロジックを構成する) で構成される。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="1d9f9-108">Outlook アドインは、[AppSource](https://appsource.microsoft.com) から入手するか、エンドユーザーまたは管理者が[サイドロード](sideload-outlook-add-ins-for-testing.md)することができます。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="1d9f9-109">Outlook アドインは、Windows 版 Outlook 固有の統合機能として以前から存在した COM アドインや VSTO アドインとは異なります。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="1d9f9-110">COM アドインとは違い、Outlook アドインのコードがユーザーのデバイスまたは Outlook クライアントに物理的にインストールされることはありません。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="1d9f9-111">Outlook のアドインの場合、Outlook はマニフェストを読み取り UI で指定したコントロールをフックした後に、HTML と JavaScript を読み込みます。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="1d9f9-112">この Web コンポーネントは、サンドボックス内のブラウザーのコンテキストですべて実行されます。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="1d9f9-113">アドインをサポートしている Outlook アイテムには、メール メッセージ、会議出席依頼、会議出席依頼の返信、会議の取り消し、予定などがあります。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="1d9f9-114">それぞれの Outlook アドインにより、アイテムの種類、ユーザーがアイテムの読み取りや作成を行うかどうかなど、使用できるコンテキストが定義されます。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

> [!NOTE]
> <span data-ttu-id="1d9f9-p104">アドインをビルドするとき、アドインを AppSource に[発行](../publish/publish.md)する予定であれば、[AppSource 検証ポリシー](/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p104">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="extension-points"></a><span data-ttu-id="1d9f9-117">拡張点</span><span class="sxs-lookup"><span data-stu-id="1d9f9-117">Extension points</span></span>

<span data-ttu-id="1d9f9-p105">拡張点は、アドインが Outlook と統合する方法です。これを行う方法は以下のとおりです。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p105">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="1d9f9-p106">アドインは、メッセージと予定のコマンド サーフェスに表示されるボタンを宣言できます。詳細は、「 [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p106">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="1d9f9-122">**リボン上の [コマンド] ボタンがあるアドイン**</span><span class="sxs-lookup"><span data-stu-id="1d9f9-122">**An add-in with command buttons on the ribbon**</span></span>

    ![アドイン コマンドの UI なし図形](../images/uiless-command-shape.png)

- <span data-ttu-id="1d9f9-p107">アドインは、メッセージおよび予定内の正規表現に一致するものや検出されたエンティティのリンクをオフにすることができます。 詳細は、「 [コンテキスト Outlook アドイン](contextual-outlook-add-ins.md)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p107">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="1d9f9-126">**強調表示されたエンティティ (アドレス) 用のコンテキスト アドイン**</span><span class="sxs-lookup"><span data-stu-id="1d9f9-126">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![カード内のコンテキスト アプリを示しています](../images/outlook-detected-entity-card.png)


> [!NOTE]
> <span data-ttu-id="1d9f9-128">[カスタム ウィンドウは廃止された](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)ため、サポートされている拡張点を使用していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-128">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.</span></span>

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="1d9f9-129">アドインで使用可能なメールボックスのアイテム</span><span class="sxs-lookup"><span data-stu-id="1d9f9-129">Mailbox items available to add-ins</span></span>

<span data-ttu-id="1d9f9-p108">Outlook アドインは、作成中や読み取り中にメッセージや予定で使用することができますが、他のアイテムの種類では使用できません。新規作成フォームまたは閲覧フォームで現在のメッセージ アイテムが次のいずれかの場合、Outlook はアドインをアクティブ化しません。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p108">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="1d9f9-p109">Information Rights Management (IRM) によって保護されているか、または保護のためにその他の方法で暗号化されている場合。デジタル署名はこれらいずれかのメカニズムに依存しているため、デジタル署名されたメッセージはその一例です。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-p109">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

- <span data-ttu-id="1d9f9-134">メッセージ クラスが IPM.Report.\* である配信レポートまたは通知 (配信レポート、配信不能レポート (NDR)、開封通知、未開封通知、遅延通知など)。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-134">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="1d9f9-135">下書きであるか (送信者が割り当てられていない)、Outlook の [下書き] フォルダーにある場合。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-135">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="1d9f9-136">別のメッセージに添付される .msg または .eml ファイルの場合。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="1d9f9-137">.msg または .eml ファイルがファイル システムから開かれた場合。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="1d9f9-138">共有メールボックス内、別のユーザーのメールボックス内、アーカイブ メールボックス内、パブリック フォルダー内。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-138">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="1d9f9-139">カスタム フォームを使用する場合。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-139">Using a custom form.</span></span>

<span data-ttu-id="1d9f9-140">既知のエンティティの文字列照合に基づいてアクティブ化されるアドインを除いて、通常、Outlook は [送信済みアイテム] フォルダーのアイテムに対して閲覧フォーム内でアドインをアクティブ化できます。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-140">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="1d9f9-141">この理由の詳細は、[Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)の「既知のエンティティに対するサポート」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-141">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-hosts"></a><span data-ttu-id="1d9f9-142">サポートされるホスト</span><span class="sxs-lookup"><span data-stu-id="1d9f9-142">Supported hosts</span></span>

<span data-ttu-id="1d9f9-143">Outlook アドインは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、オンプレミスの Exchange 2013 用 Outlook on the web 以降の各バージョン、iOS 用 Outlook、Android 用 Outlook、および Office 365 と Outlook.com の Outlook on the web でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-143">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web in Office 365 and Outlook.com.</span></span> <span data-ttu-id="1d9f9-144">最新の機能すべてが、すべての[クライアント](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)で同時にサポートされているわけではありません。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-144">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="1d9f9-145">これらの機能が各ホストでサポートされる可能性の有無については、該当する機能に関する記事や API リファレンスを参照してください。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-145">Please refer to articles and API references for those features to see which hosts they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="1d9f9-146">Outlook アドインの作成を開始する</span><span class="sxs-lookup"><span data-stu-id="1d9f9-146">Get started building Outlook add-ins</span></span>

<span data-ttu-id="1d9f9-147">Outlook アドインの作成を開始するには、次の操作を行います。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-147">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="1d9f9-148">[クイックスタート](../quickstarts/outlook-quickstart.md) - 簡単な作業ウィンドウを作成します。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-148">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="1d9f9-149">[チュートリアル](../tutorials/outlook-tutorial.md) - 新しいメッセージに GitHub gist を挿入するアドインを作成する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="1d9f9-149">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="1d9f9-150">関連項目</span><span class="sxs-lookup"><span data-stu-id="1d9f9-150">See also</span></span>

- [<span data-ttu-id="1d9f9-151">Office アドイン開発のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="1d9f9-151">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="1d9f9-152">Office アドインの設計ガイドライン</span><span class="sxs-lookup"><span data-stu-id="1d9f9-152">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1d9f9-153">Office および SharePoint アドインのライセンスを付与する</span><span class="sxs-lookup"><span data-stu-id="1d9f9-153">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="1d9f9-154">Office アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="1d9f9-154">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="1d9f9-155">AppSource と Office 内でソリューションを使用できるようにする</span><span class="sxs-lookup"><span data-stu-id="1d9f9-155">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
