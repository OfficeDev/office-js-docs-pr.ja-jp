---
title: Office アドインの開発ライフ サイクル
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 713daff9a0d16f904209f8b4561f3cf51bd9a9c9
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944283"
---
# <a name="office-add-ins-development-lifecycle"></a><span data-ttu-id="855c0-102">Office アドインの開発ライフ サイクル</span><span class="sxs-lookup"><span data-stu-id="855c0-102">Office Add-ins development lifecycle</span></span>

> [!NOTE]
> <span data-ttu-id="855c0-p101">AppSource にアドインを[公開](../publish/publish.md)し、Office エクスペリエンスで利用できるようにする予定がある場合は、[AppSource の検証ポリシー](https://docs.microsoft.com/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、定義したメソッドをサポートするすべてのプラットフォームでアドインが動作する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と [Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)のページを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="855c0-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

<span data-ttu-id="855c0-105">Office アドインの一般的な開発ライフサイクルには、次の手順が含まれます。</span><span class="sxs-lookup"><span data-stu-id="855c0-105">The typical development lifecycle of an Office Add-in includes the following steps:</span></span>


## <a name="1-decide-on-the-purpose-of-the-add-in"></a><span data-ttu-id="855c0-106">1. アドインの用途を決定する</span><span class="sxs-lookup"><span data-stu-id="855c0-106">1. Decide on the purpose of the add-in</span></span>
    
<span data-ttu-id="855c0-107">次のことを確認します。</span><span class="sxs-lookup"><span data-stu-id="855c0-107">Ask the following questions:</span></span>
    
- <span data-ttu-id="855c0-108">どのように役立つアドインですか。</span><span class="sxs-lookup"><span data-stu-id="855c0-108">How is the add-in useful?</span></span> 
        
- <span data-ttu-id="855c0-109">どのような形で顧客の生産性向上に寄与しますか。</span><span class="sxs-lookup"><span data-stu-id="855c0-109">How does it help your customers be more productive?</span></span>
        
- <span data-ttu-id="855c0-110">アドインの機能はどのようなシナリオをサポートしますか。</span><span class="sxs-lookup"><span data-stu-id="855c0-110">What scenarios does your add-in's features support?</span></span>
    
<span data-ttu-id="855c0-111">最も重要な機能とシナリオを決定し、それらに設計の重点を置きます。</span><span class="sxs-lookup"><span data-stu-id="855c0-111">Decide the most important features and scenarios and focus your design around them.</span></span> 

    
## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a><span data-ttu-id="855c0-112">2. アドインのデータおよびデータ ソースを特定する</span><span class="sxs-lookup"><span data-stu-id="855c0-112">2. Identify the data and data source for the add-in</span></span>
    
- <span data-ttu-id="855c0-113">データは、ドキュメント、ブック、プレゼンテーション、プロジェクト、または Access のブラウザーベースのデータベースに含まれるものですか。</span><span class="sxs-lookup"><span data-stu-id="855c0-113">Is the data in a document, workbook, presentation, project, or an Access browser-based database?</span></span> 
    
- <span data-ttu-id="855c0-114">データは Exchange Server や Exchange Online のメールボックスのアイテムに関するものですか。</span><span class="sxs-lookup"><span data-stu-id="855c0-114">Is the data about an item or items in an Exchange Server or Exchange Online mailbox?</span></span> 
    
- <span data-ttu-id="855c0-115">データは Web サービスなどの外部ソースからのものですか。</span><span class="sxs-lookup"><span data-stu-id="855c0-115">Is the data from an external source such as a web service?</span></span>

    
## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a><span data-ttu-id="855c0-116">3. アドインの種類を判断し、アドインの目的に最も合致する Office ホスト アプリケーションを特定する</span><span class="sxs-lookup"><span data-stu-id="855c0-116">3. Identify the type of add-in and Office host applications that best support the purpose of the add-in</span></span>
    
<span data-ttu-id="855c0-117">次のことを考慮してシナリオを特定します。</span><span class="sxs-lookup"><span data-stu-id="855c0-117">Consider the following to identify the scenarios:</span></span>
    
- <span data-ttu-id="855c0-p102">ユーザーはドキュメントや Access ブラウザーベースのデータベースの内容を充実させるためにアドインを使用しますか。その場合は、**コンテンツ アドイン**の作成を検討します。</span><span class="sxs-lookup"><span data-stu-id="855c0-p102">Will customers use the add-in to enrich the content of a document or Access browser-based database? If so, you may want to consider creating a **content add-in**.</span></span> 
    
- <span data-ttu-id="855c0-p103">ユーザーはメール メッセージや予定を表示または作成するときにアドインを使いますか。現在のコンテキストに従ってアドインを公開できることが重要ですか。デスクトップだけでなくタブレットやスマートフォンでもアドインを使用できるようにすることが優先されますか。</span><span class="sxs-lookup"><span data-stu-id="855c0-p103">Will customers use the add-in while viewing or composing an email message or appointment? Is being able to expose the add-in according to the current context important? Is making the add-in available on not just the desktop, but also on tablets and phones a priority?</span></span>
    
    <span data-ttu-id="855c0-p104">これらの質問のいずれかに「はい」と答えた場合は、**Outlook アドイン**の作成を検討します。その後、アドインをトリガーするコンテキストを明らかにします (作成フォーム、特定のメッセージ タイプ、添付ファイル、アドレス、タスクのヒント、または会議提案の存在、メールや予定の内容に特定の文字列パターンなど)。</span><span class="sxs-lookup"><span data-stu-id="855c0-p104">If you answer yes to any of these questions, consider creating an **Outlook add-in**. Identify the context that will trigger your add-in (for example, the user being in a compose form, specific message types, the presence of an attachment, address, task suggestion, or meeting suggestion, or certain string patterns in the contents of an email or appointment).</span></span> 
        
    <span data-ttu-id="855c0-125">Outlook アドインのコンテキストによるアクティブ化方法については、「[Outlook アドインのアクティブ化ルール](https://docs.microsoft.com/outlook/add-ins/activation-rules)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855c0-125">To find out how you can contextually activate the Outlook add-in, see [Activation rules for Outlook add-ins](https://docs.microsoft.com/outlook/add-ins/activation-rules).</span></span> 
    
- <span data-ttu-id="855c0-p105">ユーザーはドキュメントの表示または作成エクスペリエンスを向上するためにアドインを使用しますか。その場合は、**作業ウィンドウ アドイン**の作成を検討します。</span><span class="sxs-lookup"><span data-stu-id="855c0-p105">Will customers use the add-in to enhance the viewing or authoring experience of a document? If so, you may want to consider creating a **task pane add-in**.</span></span> 

<span data-ttu-id="855c0-128">Office アプリケーションと、それが動作しているプラットフォーム (Windows、Mac、Web、モバイル) では、特定のアドイン API のサポートが異なる場合があります。</span><span class="sxs-lookup"><span data-stu-id="855c0-128">Support for certain Add-in APIs may differ between Office applications and the platform they are running on (Windows, Mac, Web, Mobile).</span></span> <span data-ttu-id="855c0-129">クライアントとプラットフォームによる現在の API 対応を確認するには、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855c0-129">To see the current API coverage by client and platform, please see our [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>  

    
## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a><span data-ttu-id="855c0-130">4. アドインのユーザー エクスペリエンスとユーザー インターフェイスを設計および実装する</span><span class="sxs-lookup"><span data-stu-id="855c0-130">4. Design and implement the user experience and user interface for the add-in</span></span>
    
<span data-ttu-id="855c0-p107">一貫性があり、習得しやすく、主要なシナリオを数ステップの手順で完了できるような、迅速で円滑なユーザー エクスペリエンスを設計します。アドインの目的によっては、サードパーティの API や Web サービスを利用します。</span><span class="sxs-lookup"><span data-stu-id="855c0-p107">Design a fast and fluid user experience that is consistent, easy to learn, with primary scenarios that require only a few steps to complete. Depending on the purpose of the add-in, make use of third-party APIs or web services.</span></span>
    
<span data-ttu-id="855c0-133">さまざまな Web 開発ツールを選択でき、HTML と JavaScript を使用してユーザー インターフェイスを実装できます。</span><span class="sxs-lookup"><span data-stu-id="855c0-133">You can choose from a variety of web development tools and use HTML and JavaScript to implement the user interface.</span></span>

    
## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a><span data-ttu-id="855c0-134">5. Office アドイン マニフェスト スキーマに基づく XML マニフェスト ファイルを作成する</span><span class="sxs-lookup"><span data-stu-id="855c0-134">5. Create an XML manifest file based on the Office Add-ins manifest schema</span></span>
    
<span data-ttu-id="855c0-135">XML マニフェストを作成します。この中に、アドインとその要件を識別する情報を記述します。また、アドインが使用する HTML ファイル、JavaScript ファイル、および CSS ファイルの場所を指定し、アドインの種類によっては既定のサイズとアクセス許可も指定します。</span><span class="sxs-lookup"><span data-stu-id="855c0-135">Create an XML manifest to identify the add-in and its requirements, specify the locations of the HTML and any JavaScript and CSS files that the add-in uses, and depending on the type of the add-in, the default size and permissions.</span></span>
    
<span data-ttu-id="855c0-p108">Outlook アドインの場合は、現在のメッセージまたは予定に基づいてコンテキストを指定できます。そのコンテキストのもとでアドインは意味を持ち、Outlook の UI で使用できるようになります。また、アドインがサポートするデバイスを決定することもできます。マニフェストで、コンテキストをアクティブ化ルールとして指定し、サポート対象デバイスを指定します。</span><span class="sxs-lookup"><span data-stu-id="855c0-p108">For Outlook add-ins, you can specify the context, based on the current message or appointment, under which your add-in is relevant and you would like Outlook to make available in the UI. You can also decide which devices you want the add-in to support. In the manifest, specify the context as activation rules and the supported devices.</span></span>
    

## <a name="6-install-and-test-the-add-in"></a><span data-ttu-id="855c0-139">6. アドインをインストールおよびテストする</span><span class="sxs-lookup"><span data-stu-id="855c0-139">6. Install and test the add-in</span></span>
    
<span data-ttu-id="855c0-p109">アドインのマニフェスト ファイルで指定した Web サーバーに、HTML ファイル、JavaScript ファイル、CSS ファイルを配置します。アドインをインストールする手順は、アドインの種類によって異なります。詳細については、「[テスト用に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855c0-p109">Place the HTML files and any JavaScript and CSS files on the web servers that are specified in the add-in manifest file. The process to install an add-in depends on the type of the add-in. For details, see [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>
    
<span data-ttu-id="855c0-p110">Outlook アドインの場合、Exchange メールボックスにインストールし、Exchange 管理センター (EAC) でアドインのマニフェスト ファイルの場所を指定します。詳細については、「[テスト用に Outlook アドインを展開してインストールする](https://docs.microsoft.com/outlook/add-ins/testing-and-tips)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855c0-p110">For Outlook add-ins, install it in an Exchange mailbox, and specify the location of the add-in manifest file in the Exchange Admin Center (EAC). For more information, see [Deploy and install Outlook add-ins for testing](https://docs.microsoft.com/outlook/add-ins/testing-and-tips).</span></span>

    
## <a name="7-publish-the-add-in"></a><span data-ttu-id="855c0-145">7. アドインを発行する</span><span class="sxs-lookup"><span data-stu-id="855c0-145">7. Publish the add-in</span></span>
    
<span data-ttu-id="855c0-p111">アドインを AppSource に送信できます。お客様はそこからアドインをインストールできます。さらに、作業ウィンドウおよびコンテンツのアドインを SharePoint 上のプライベート フォルダー アドイン カタログまたは共有ネットワーク フォルダーに発行することが可能で、組織の Exchange サーバーに Outlook アドインを直接展開できます。詳細については、「[Office アドインを発行する](../publish/publish.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="855c0-p111">You can submit the add-in to AppSource, from which customers can install the add-in. In addition, you can publish task pane and content add-ins to a private folder add-in catalog on SharePoint or to a shared network folder, and you can deploy an Outlook add-in directly on an Exchange server for your organization. For details, see [Publish your Office Add-in](../publish/publish.md).</span></span>
    
    
## <a name="8-maintain-the-add-in"></a><span data-ttu-id="855c0-149">8. アドインをメンテナンスする</span><span class="sxs-lookup"><span data-stu-id="855c0-149">8. Maintain the add-in</span></span>
    
<span data-ttu-id="855c0-150">アドインから Web サービスを呼び出していて、アドインの公開後に Web サービスを更新する場合、アドインを再発行する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="855c0-150">If your add-in calls a web service, and if you make updates to the web service after publishing the add-in, you do not have to republish the add-in.</span></span> <span data-ttu-id="855c0-151">ただし、アドイン マニフェスト、スクリーンショット、アイコン、HTML、JavaScript のファイルなど、アドインに送信したアイテムやデータを変更する場合は、アドインを再発行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="855c0-151">However, if you change any items or data you submitted for your add-in, such as the add-in manifest, screenshots, icons, HTML or JavaScript files, you will need to republish the add-in.</span></span> 
    
<span data-ttu-id="855c0-p113">具体的には、AppSource にアドインを発行した場合は、AppSource が変更を実装できるようにアドインを再送信する必要があります。アドインと一緒に、新しいバージョン番号を含む更新されたアドイン マニフェストを再送信する必要があります。また、新しいマニフェストのバージョン番号と一致するように、送信フォームのアドイン バージョン番号を更新する必要があります。Outlook アドインの場合は、[ID](https://docs.microsoft.com/javascript/office/manifest/id?view=office-js) 要素にアドイン マニフェストの異なる UUID が含まれることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="855c0-p113">In particular, if you have published the add-in to AppSource, you'll need to resubmit your add-in so that AppSource can implement those changes. You must resubmit your add-in with an updated add-in manifest that includes a new version number. You must also make sure to update the add-in version number in the submission form to match the new manifest's version number. For Outlook add-ins, you should make sure the [Id](https://docs.microsoft.com/javascript/office/manifest/id?view=office-js) element contains a different UUID in the add-in manifest.</span></span>
    
