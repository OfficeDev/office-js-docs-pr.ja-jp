---
title: Word アドインの概要
description: Word アドインの基礎の説明
ms.date: 03/18/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: f176f8ed190642cf047686f78bc2407f686bdf60
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891034"
---
# <a name="word-add-ins-overview"></a><span data-ttu-id="7452b-103">Word アドインの概要</span><span class="sxs-lookup"><span data-stu-id="7452b-103">Word add-ins overview</span></span>

<span data-ttu-id="7452b-p101">Word の機能を拡張するソリューション (たとえば、ドキュメントの自動アセンブリや、他のデータ ソースからの Word 文書のデータへのバインドやアクセスを可能にするソリューション) を作成したい場合は、Word JavaScript API と Office JavaScript API を含む Office アドイン プラットフォームを使用して、Windows デスクトップ、Mac、またはクラウドで実行する Word クライアントを拡張できます。</span><span class="sxs-lookup"><span data-stu-id="7452b-p101">Do you want to create a solution that extends the functionality of Word? For example, one that involves automated document assembly? Or a solution that binds to and accesses data in a Word document from other data sources? You can use the Office Add-ins platform, which includes the Word JavaScript API and the Office JavaScript API, to extend Word clients running on a Windows desktop, on a Mac, or in the cloud.</span></span>

<span data-ttu-id="7452b-p102">Word のアドインは、[Office アドイン プラットフォーム](../overview/office-add-ins.md)にある数多くの開発オプションのひとつです。アドイン コマンドを使用して、Word の UI を拡張したり、Word 文書のコンテンツと対話する JavaScript を実行する作業ウィンドウを起動したりすることができます。ブラウザーで実行できるあらゆるコードは、Word アドインで実行できます。Word 文書のコンテンツと対話するアドインは、Word オブジェクトを操作し、オブジェクトの状態を同期する要求を作成します。</span><span class="sxs-lookup"><span data-stu-id="7452b-p102">Word add-ins are one of the many development options that you have on the [Office Add-ins platform](../overview/office-add-ins.md). You can use add-in commands to extend the Word UI and launch task panes that run JavaScript that interacts with the content in a Word document. Any code that you can run in a browser can run in a Word add-in. Add-ins that interact with content in a Word document create requests to act on Word objects and synchronize object state.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="7452b-112">次の図は、作業ウィンドウで実行される Word アドインの例を示します。</span><span class="sxs-lookup"><span data-stu-id="7452b-112">The following figure shows an example of a Word add-in that runs in a task pane.</span></span>

<span data-ttu-id="7452b-113">*図 1. Word の作業ウィンドウで実行されているアドイン*</span><span class="sxs-lookup"><span data-stu-id="7452b-113">*Figure 1. Add-in running in a task pane in Word*</span></span>

![Word の作業ウィンドウで実行されているアドイン](../images/word-add-in-show-host-client.png)

<span data-ttu-id="7452b-p103">Word アドイン (1) は、Word 文書 (2) に要求を送信し、JavaScript を使用して段落オブジェクトにアクセスして段落を更新、削除、または移動することができます。たとえば、次のコードは、その段落に新しい文を追加する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="7452b-p103">The Word add-in (1) can send requests to the Word document (2) and can use JavaScript to access the paragraph object and update, delete, or move the paragraph. For example, the following code shows how to append a new sentence to that paragraph.</span></span>

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

<span data-ttu-id="7452b-p104">ASP.NET、NodeJS、Python などの任意の Web サーバー テクノロジを使用して、Word アドインをホストさせることができます。お気に入りのクライアント側のフレームワーク (Ember、Backbone、Angular、React) を使用するか VanillaJS を引き続き使用してソリューションを開発し、Azure のようなサービスを使用してアプリケーションを[認証](../develop/overview-authn-authz.md)し、ホストできます。</span><span class="sxs-lookup"><span data-stu-id="7452b-p104">You can use any web server technology to host your Word add-in, such as ASP.NET, NodeJS, or Python. Use your favorite client-side framework -- Ember, Backbone, Angular, React -- or stick with VanillaJS to develop your solution, and you can use services like Azure to [authenticate](../develop/overview-authn-authz.md) and host your application.</span></span>

<span data-ttu-id="7452b-p105">Word JavaScript API を使用すると、アプリケーションから Word 文書内にあるオブジェクトやメタデータにアクセスできます。これらの API を使用して、以下をターゲットとするアドインを作成できます。</span><span class="sxs-lookup"><span data-stu-id="7452b-p105">The Word JavaScript APIs give your application access to the objects and metadata found in a Word document. You can use these APIs to create add-ins that target:</span></span>

* <span data-ttu-id="7452b-121">Windows での Word 2013 以降</span><span class="sxs-lookup"><span data-stu-id="7452b-121">Word 2013 or later on Windows</span></span>
* <span data-ttu-id="7452b-122">Word on the web</span><span class="sxs-lookup"><span data-stu-id="7452b-122">Word on the web</span></span>
* <span data-ttu-id="7452b-123">Mac on Word 2016 以降</span><span class="sxs-lookup"><span data-stu-id="7452b-123">Word 2016 or later on Mac</span></span>
* <span data-ttu-id="7452b-124">Word on iPad</span><span class="sxs-lookup"><span data-stu-id="7452b-124">Word on iPad</span></span>

<span data-ttu-id="7452b-p106">アドインを 1 回作成すれば、それをプラットフォームの異なるすべてのバージョンの Word で実行できます。詳細については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7452b-p106">Write your add-in once, and it will run in all versions of Word across multiple platforms. For details, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="javascript-apis-for-word"></a><span data-ttu-id="7452b-127">Word 用 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="7452b-127">JavaScript APIs for Word</span></span>

<span data-ttu-id="7452b-128">2 セットの JavaScript API を使用して、Word 文書のオブジェクトおよびメタデータと対話できます。</span><span class="sxs-lookup"><span data-stu-id="7452b-128">You can use two sets of JavaScript APIs to interact with the objects and metadata in a Word document.</span></span> <span data-ttu-id="7452b-129">1 つ目は、Office 2013 で導入された[共通 API](/javascript/api/office) です。</span><span class="sxs-lookup"><span data-stu-id="7452b-129">The first is the [Common API](/javascript/api/office), which was introduced in Office 2013.</span></span> <span data-ttu-id="7452b-130">2 つ以上の Office クライアントでホストされているアドインで、共通 API の多くのオブジェクトを使用することができます。</span><span class="sxs-lookup"><span data-stu-id="7452b-130">Many of the objects in the Common API can be used in add-ins hosted by two or more Office clients.</span></span> <span data-ttu-id="7452b-131">この API は、広範囲にわたってコールバックを使用します。</span><span class="sxs-lookup"><span data-stu-id="7452b-131">This API uses callbacks extensively.</span></span>

<span data-ttu-id="7452b-p108">2 つ目は、[Word JavaScript API](/javascript/api/word) です。これは、Mac と Windows の Word 2016 を対象とする Word アドインを作成するために使用できる、厳密に型指定されたオブジェクト モデルです。このオブジェクト モデルは promise を使用し、[本文](/javascript/api/word/word.body)、[コンテンツ コントロール](/javascript/api/word/word.contentcontrol)、[インライン画像](/javascript/api/word/word.inlinepicture)、および[段落](/javascript/api/word/word.paragraph)などの Word 固有のオブジェクトへのアクセスを提供します。Word JavaScript API には、IDE 内のコード ヒントを取得できるように、TypeScript の定義と vsdoc ファイルが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7452b-p108">The second is the [Word JavaScript API](/javascript/api/word). This is a strongly-typed object model that you can use to create Word add-ins that target Word 2016 on Mac and Windows. This object model uses promises, and provides access to Word-specific objects like [body](/javascript/api/word/word.body), [content controls](/javascript/api/word/word.contentcontrol), [inline pictures](/javascript/api/word/word.inlinepicture), and [paragraphs](/javascript/api/word/word.paragraph). The Word JavaScript API includes TypeScript definitions and vsdoc files so that you can get code hints in your IDE.</span></span>

<span data-ttu-id="7452b-p109">現在、Word のすべてのクライアントは共有の Office JavaScript API をサポートし、ほとんどのクライアントは Word JavaScript API をサポートします。サポート対象のクライアントの詳細については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7452b-p109">Currently, all Word clients support the shared Office JavaScript API, and most clients support the Word JavaScript API. For details about supported clients, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

<span data-ttu-id="7452b-p110">Word JavaScript API のオブジェクト モデルはより簡単に使用できるため、Word JavaScript APから始めることをお勧めします。次のような必要がある場合は、Word JavaScript API を使用します。</span><span class="sxs-lookup"><span data-stu-id="7452b-p110">We recommend that you start with the Word JavaScript API because the object model is easier to use. Use the Word JavaScript API if you need to:</span></span>

* <span data-ttu-id="7452b-140">Word 文書内のオブジェクトにアクセスする。</span><span class="sxs-lookup"><span data-stu-id="7452b-140">Access the objects in a Word document.</span></span>

<span data-ttu-id="7452b-141">次のような必要がある場合は、共有の Office JavaScript API を使用します。</span><span class="sxs-lookup"><span data-stu-id="7452b-141">Use the shared Office JavaScript API when you need to:</span></span>

* <span data-ttu-id="7452b-142">Word 2013 を対象とする。</span><span class="sxs-lookup"><span data-stu-id="7452b-142">Target Word 2013.</span></span>
* <span data-ttu-id="7452b-143">アプリケーションの初期アクションを実行する。</span><span class="sxs-lookup"><span data-stu-id="7452b-143">Perform initial actions for the application.</span></span>
* <span data-ttu-id="7452b-144">サポートされている要件セットを確認する。</span><span class="sxs-lookup"><span data-stu-id="7452b-144">Check the supported requirement set.</span></span>
* <span data-ttu-id="7452b-145">メタデータ、設定、およびドキュメントの環境情報にアクセスする。</span><span class="sxs-lookup"><span data-stu-id="7452b-145">Access metadata, settings, and environmental information for the document.</span></span>
* <span data-ttu-id="7452b-146">ドキュメント内のセクションにバインドし、イベントをキャプチャする。</span><span class="sxs-lookup"><span data-stu-id="7452b-146">Bind to sections in a document and capture events.</span></span>
* <span data-ttu-id="7452b-147">カスタム XML パーツを使用する。</span><span class="sxs-lookup"><span data-stu-id="7452b-147">Use custom XML parts.</span></span>
* <span data-ttu-id="7452b-148">ダイアログ ボックスを開く。</span><span class="sxs-lookup"><span data-stu-id="7452b-148">Open a dialog box.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7452b-149">次の手順</span><span class="sxs-lookup"><span data-stu-id="7452b-149">Next steps</span></span>

<span data-ttu-id="7452b-p111">最初の Word アドインを作成する準備ができたら「[最初の Word アドインをビルドする](word-add-ins.md)」を参照してください。[アドインのマニフェスト](../develop/add-in-manifests.md) を使用して、アドインがホストされている場所や表示方法の説明、アクセス許可およびその他の情報の定義を行います。</span><span class="sxs-lookup"><span data-stu-id="7452b-p111">Ready to create your first Word add-in? See [Build your first Word add-in](word-add-ins.md). Use the [add-in manifest](../develop/add-in-manifests.md) to describe where your add-in is hosted, how it is displayed, and define permissions and other information.</span></span>

<span data-ttu-id="7452b-153">ユーザーにとって魅力的なエクスペリエンスを提供する世界クラスの Word アドインを設計する方法の詳細については、「[設計のガイドライン](../design/add-in-design.md)」と「[ベスト プラクティス](../concepts/add-in-development-best-practices.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7452b-153">To learn more about how to design a world class Word add-in that creates a compelling experience for your users, see [Design guidelines](../design/add-in-design.md) and [Best practices](../concepts/add-in-development-best-practices.md).</span></span>

<span data-ttu-id="7452b-154">アドインを作成したら、ネットワーク共有、アプリ カタログ、または AppSource に[発行](../publish/publish.md)できます。</span><span class="sxs-lookup"><span data-stu-id="7452b-154">After you develop your add-in, you can [publish](../publish/publish.md) it to a network share, an app catalog, or AppSource.</span></span>

## <a name="see-also"></a><span data-ttu-id="7452b-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="7452b-155">See also</span></span>

* [<span data-ttu-id="7452b-156">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="7452b-156">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="7452b-157">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="7452b-157">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="7452b-158">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7452b-158">Word JavaScript API reference</span></span>](../reference/overview/word-add-ins-reference-overview.md)
