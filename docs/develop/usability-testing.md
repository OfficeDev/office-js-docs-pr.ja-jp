---
title: Office アドインのユーザビリティ テスト
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 4b21af2502c9357e8a7d2c953cd5182833577ac9
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925452"
---
# <a name="usability-testing-for-office-add-ins"></a><span data-ttu-id="a8068-102">Office アドインのユーザビリティ テスト</span><span class="sxs-lookup"><span data-stu-id="a8068-102">Usability testing for Office Add-ins</span></span>

<span data-ttu-id="a8068-p101">ユーザーの動作を考慮してデザインされたアドインは優れています。デザインを決定する際には自分の先入感が影響を及ぼすので、実際のユーザーを用いてデザインをテストし、アドインが顧客にとって有用かどうかを確認するのは重要です。</span><span class="sxs-lookup"><span data-stu-id="a8068-p101">A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.</span></span> 

<span data-ttu-id="a8068-p102">ユーザビリティ テストは、さまざまな方法で実行できます。多くのアドインの開発者にとっては、リモートで、モデレートせずにユーザビリティを検討することが、時間とコストの点で最も効率的です。このことを簡単に行える、人気のテスト サービスがいくつかあります。次に、例を示します。</span><span class="sxs-lookup"><span data-stu-id="a8068-p102">You can run usability tests in different ways. For many add-in developers, remote, unmoderated usability studies are the most time and cost effective. Several popular testing services make this easy; the following are some examples:</span></span> 

 - [<span data-ttu-id="a8068-108">UserTesting.com</span><span class="sxs-lookup"><span data-stu-id="a8068-108">UserTesting.com</span></span>](https://www.UserTesting.com)
 - [<span data-ttu-id="a8068-109">Optimalworkshop.com</span><span class="sxs-lookup"><span data-stu-id="a8068-109">Optimalworkshop.com</span></span>](https://www.Optimalworkshop.com)
 - [<span data-ttu-id="a8068-110">Userzoom.com</span><span class="sxs-lookup"><span data-stu-id="a8068-110">Userzoom.com</span></span>](https://www.Userzoom.com)

<span data-ttu-id="a8068-111">これらのテスト サービスは、テスト計画の作成を効率化し、参加者を探したりテストをモデレートしたりする必要をなくすのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="a8068-111">These testing services help you to streamline test plan creation and remove the need to seek out participants or moderate the tests.</span></span> 

<span data-ttu-id="a8068-p103">デザインに関するほとんどのユーザビリティの問題を見出すのに必要な参加者は 5 人だけです。製品がユーザー中心になっていることを確認するには、開発サイクル全体にわたって小規模なテストを定期的に組み込んでください。</span><span class="sxs-lookup"><span data-stu-id="a8068-p103">You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.</span></span>

> [!NOTE]
> <span data-ttu-id="a8068-p104">複数のプラットフォームにまたがってアドインのユーザビリティをテストすることをお勧めします。[AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store) にアドインを公開するには、[定義したメソッドをサポートするプラットフォーム](../overview/office-add-in-availability.md)すべてでそのアドインが作動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8068-p104">We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store), it must work on all [platforms that support the methods that you define](../overview/office-add-in-availability.md).</span></span>

## <a name="1---sign-up-for-a-testing-service"></a><span data-ttu-id="a8068-116">1.   テスト サービスにサインアップする</span><span class="sxs-lookup"><span data-stu-id="a8068-116">1.   Sign up for a testing service</span></span>

<span data-ttu-id="a8068-117">詳細については、「[モデレートされていないリモート ユーザー テスト用のオンライン ツールの選択](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-117">For more information, see [Selecting an Online Tool for Unmoderated Remote User Testing.](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)</span></span>

## <a name="2-develop-your-research-questions"></a><span data-ttu-id="a8068-118">2.調査での質問項目を設定する</span><span class="sxs-lookup"><span data-stu-id="a8068-118">2. Develop your research questions</span></span>
 
<span data-ttu-id="a8068-p105">調査での質問項目を設定することにより、調査の目的を定義し、テストの計画を導いていることになります。質問項目は、募集する参加者や実行するタスクを特定するのに役立ちます。調査での質問項目は、可能な限り具体的に設定してください。広範な質問に回答するように努めることもできます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p105">Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.</span></span>
 
<span data-ttu-id="a8068-123">次に、調査での質問項目の設定例をいくつか示します。</span><span class="sxs-lookup"><span data-stu-id="a8068-123">The following are some examples of research questions:</span></span>
  
<span data-ttu-id="a8068-124">**具体的な質問**</span><span class="sxs-lookup"><span data-stu-id="a8068-124">**Specific**</span></span>  

 - <span data-ttu-id="a8068-125">ユーザーは、ランディング ページ上の "無料試用版" リンクに気が付きますか。</span><span class="sxs-lookup"><span data-stu-id="a8068-125">Do users notice the "free trial" link on the landing page?</span></span>
 - <span data-ttu-id="a8068-126">ユーザーがアドインから自分のドキュメントにコンテンツを挿入する際に、ドキュメント内の挿入場所がわかるでしょうか。</span><span class="sxs-lookup"><span data-stu-id="a8068-126">When users insert content from the add-in to their document, do they understand where in the document it is inserted?</span></span>

<span data-ttu-id="a8068-127">**広範な質問**</span><span class="sxs-lookup"><span data-stu-id="a8068-127">**Broad**</span></span>  

 - <span data-ttu-id="a8068-128">ユーザーにとってアドインの最大の懸案事項は何ですか。</span><span class="sxs-lookup"><span data-stu-id="a8068-128">What are the biggest pain points for the user in our add-in?</span></span>
 - <span data-ttu-id="a8068-129">ユーザーは、コマンド バー内のアイコンの意味を、クリックする前に理解していますか。</span><span class="sxs-lookup"><span data-stu-id="a8068-129">Do users understand the meaning of the icons in our command bar, before they click on them?</span></span>
 - <span data-ttu-id="a8068-130">ユーザーは設定メニューを簡単に検索できますか。</span><span class="sxs-lookup"><span data-stu-id="a8068-130">Can users easily find the settings menu?</span></span>

<span data-ttu-id="a8068-p106">アドインを見つけることからインストールして使用することに至るまで、ユーザー体験全体に関するデータを取得することは重要です。調査での質問項目が、アドイン ユーザー体験の次の側面に対応しているか検討してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-p106">It’s important to get data on the entire user journey – from discovering your add-in, to installing and using it. Consider research questions that address the following aspects of the add-in user experience:</span></span>
 
 - <span data-ttu-id="a8068-133">AppSource 内でのアドインの検索</span><span class="sxs-lookup"><span data-stu-id="a8068-133">Finding your add-in in AppSource</span></span>
 - <span data-ttu-id="a8068-134">アドインのインストールの選択</span><span class="sxs-lookup"><span data-stu-id="a8068-134">Choosing to install your add-in</span></span>
 - <span data-ttu-id="a8068-135">最初の実行エクスペリエンス</span><span class="sxs-lookup"><span data-stu-id="a8068-135">First run experience</span></span>
 - <span data-ttu-id="a8068-136">リボン コマンド</span><span class="sxs-lookup"><span data-stu-id="a8068-136">Ribbon commands</span></span>
 - <span data-ttu-id="a8068-137">アドイン UI</span><span class="sxs-lookup"><span data-stu-id="a8068-137">Add-in UI</span></span>
 - <span data-ttu-id="a8068-138">アドインが Office アプリケーションのドキュメント領域と相互作用する方法</span><span class="sxs-lookup"><span data-stu-id="a8068-138">How the add-in interacts with the document space of the Office application</span></span>
 - <span data-ttu-id="a8068-139">ユーザーがコンテンツ挿入フローを制御できる程度</span><span class="sxs-lookup"><span data-stu-id="a8068-139">How much control the user has over any content insertion flows</span></span>

<span data-ttu-id="a8068-140">詳細については、[効果的な質問の作成](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-140">For more information, see [Writing Effective Questions.](http://help.usertesting.com/customer/en/portal/articles/2077663-writing-effective-questions)</span></span>
 
## <a name="3-identify-participants-to-target"></a><span data-ttu-id="a8068-141">3.ターゲットとする参加者を特定する</span><span class="sxs-lookup"><span data-stu-id="a8068-141">3. Identify participants to target</span></span>
 
<span data-ttu-id="a8068-p107">リモートのテスト サービスで、テストの参加者の多くの特性を制御できます。ターゲットとするユーザーの種類を慎重に検討してください。データ収集の初期段階では、幅広く参加者を募集して、ユーザビリティの問題をより明確に識別する方が良い場合があります。後の段階では、上級の Office ユーザー、特定の職業、特定の年齢範囲などのグループをターゲットとするよう選択することもできます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p107">Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.</span></span>
 
## <a name="4-create-the-participant-screener"></a><span data-ttu-id="a8068-146">4.参加者のスクリーナーを作成する</span><span class="sxs-lookup"><span data-stu-id="a8068-146">4. Create the participant screener</span></span>
 
<span data-ttu-id="a8068-p108">スクリーナーとは、テスト用に選別するために、テストの参加予定者に提示する一連の質問と要件のことです。UserTesting.com などのサービスの参加者は、お金目的でテストの条件を満たそうとしていることに留意してください。特定のユーザーをテストから除外しようとしている場合は、スクリーナーに落とし穴のある質問を組み込むことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a8068-p108">The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test.</span></span> 
 
<span data-ttu-id="a8068-150">たとえば、GitHub に精通している参加者を見つけようとしている場合に、身元を偽っている可能性があるユーザーを除外するには、考えられる回答の一覧に偽の回答を組み込みます。</span><span class="sxs-lookup"><span data-stu-id="a8068-150">For example, if you want to find participants who are familiar with GitHub, to filter out users who might misrepresent themselves, include fakes in the list of possible answers.</span></span>

<span data-ttu-id="a8068-151">**次のソース コード リポジトリのうちどれに精通していますか。**</span><span class="sxs-lookup"><span data-stu-id="a8068-151">**Which of the following source code repositories are you familiar with?**</span></span>  
 <span data-ttu-id="a8068-p109">a. SourceShelf  *[拒否]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p109">a. SourceShelf  [*Reject*]</span></span>  
 <span data-ttu-id="a8068-p110">b. CodeContainer  *[拒否]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p110">b. CodeContainer  [*Reject*]</span></span>  
 <span data-ttu-id="a8068-p111">c. GitHub  *[必ず選択]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p111">c. GitHub  [*Must select*]</span></span>  
 <span data-ttu-id="a8068-p112">d. BitBucket  *[選択する可能性あり]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p112">d. BitBucket  [*May select*]</span></span>  
 <span data-ttu-id="a8068-p113">e. CloudForge  *[選択する可能性あり]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p113">e. CloudForge  [*May select*]</span></span>  

<span data-ttu-id="a8068-162">アドインの Live ビルドのテストを計画している場合は、次の質問により、このテストを行えるユーザーを選別できます。</span><span class="sxs-lookup"><span data-stu-id="a8068-162">If you are planning to test a live build of your add-in, the following questions can screen for users who will be able to do this.</span></span> 

<span data-ttu-id="a8068-163">**このテストには、Microsoft PowerPoint 2016 が必要です。PowerPoint 2016 がありますか。**</span><span class="sxs-lookup"><span data-stu-id="a8068-163">**This test requires you to have Microsoft PowerPoint 2016. Do you have PowerPoint 2016?**</span></span>  
 <span data-ttu-id="a8068-p114">a. はい *[必ず選択]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p114">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="a8068-p115">b. いいえ *[拒否]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p115">b. No [*Reject*]</span></span>  
 <span data-ttu-id="a8068-p116">c. わからない *[拒否]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p116">c. I don’t know [*Reject*]</span></span>  

<span data-ttu-id="a8068-170">**このテストでは、PowerPoint 2016 の無料のアドインをインストールし、これを使用するために無料のアカウントを作成する必要があります。アドインをインストールし、無料のアカウントを作成する予定がありますか。**</span><span class="sxs-lookup"><span data-stu-id="a8068-170">**This test requires you to install a free add-in for PowerPoint 2016, and create a free account to use it. Are you willing to install an add-in and create a free account?**</span></span>  
 <span data-ttu-id="a8068-p117">a. はい *[必ず選択]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p117">a. Yes [*Must select*]</span></span>  
 <span data-ttu-id="a8068-p118">b. いいえ *[拒否]*</span><span class="sxs-lookup"><span data-stu-id="a8068-p118">b. No [*Reject*]</span></span>  

<span data-ttu-id="a8068-175">詳細については、[スクリーナーの質問のベスト プラクティス](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-175">For more information, see [Screener Questions Best Practices.](http://help.usertesting.com/customer/en/portal/articles/2077835-screener-question-best-practices)</span></span>
 
## <a name="5-create-tasks-and-questions-for-participants"></a><span data-ttu-id="a8068-176">5.参加者のタスクと質問を作成する</span><span class="sxs-lookup"><span data-stu-id="a8068-176">5. Create tasks and questions for participants</span></span>
 
<span data-ttu-id="a8068-p119">参加者のタスクと質問の数を制限できるように、テストしようとしている項目の優先度の設定を試みてください。一定の時間だけ参加者に利益になるサービスもあるので、超過しないように確認することもできます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p119">Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.</span></span>

<span data-ttu-id="a8068-p120">可能な限り、参加者の動作について質問するのではなく観察するよう試みてください。動作について質問する必要がある場合、特定の状況で参加者が行うであろうと予想することではなく、参加者が過去に行ってきたことについて質問してください。その方が、信頼性の高い結果が得られる傾向があります。</span><span class="sxs-lookup"><span data-stu-id="a8068-p120">Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.</span></span>
 
<span data-ttu-id="a8068-p121">モデレートされていないテストの主な課題は、参加者がタスクとシナリオを確実に理解することです。指示は*明確で簡潔な*ものである必要があります。混乱する可能性がある場合には、必ず混乱する人がいます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p121">The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.</span></span> 

<span data-ttu-id="a8068-p122">テスト中の特定の時点で、ユーザーは想定している画面上にいるとは限らないことに注意してください。次のタスクを開始するにはどの画面上にいる必要があるかユーザーに伝えることを検討してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-p122">Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.</span></span> 

<span data-ttu-id="a8068-187">詳細については、「[Writing Great Tasks (優れたタスクの作成)](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="a8068-187">For more information, see [Writing Great Tasks.](http://help.usertesting.com/customer/en/portal/articles/2077824-writing-great-tasks)</span></span>

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a><span data-ttu-id="a8068-188">6.タスクや質問に対応するプロトタイプを作成する</span><span class="sxs-lookup"><span data-stu-id="a8068-188">6. Create a prototype to match the tasks and questions</span></span>
 
<span data-ttu-id="a8068-p123">Live アドインをテストするか、プロトタイプをテストすることができます。Live アドインをテストしようとしている場合は、Office 2016 があり、アドインのインストールを予定しており、アカウントのサインアップを予定している (ログオン資格情報を提供しない場合) 参加者を選別する必要があることに留意してください。後で、アドインを正常にインストールしたか確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a8068-p123">You can either test your live add-in, or you can test a prototype. Keep in mind that if you want to test the live add-in, you need to screen for participants that have Office 2016, are willing to install the add-in, and are willing to sign up for an account (unless you have logon credentials to provide them.) You'll then need to make sure that they successfully install your add-in.</span></span> 

<span data-ttu-id="a8068-p124">平均すると、ユーザーがアドインのインストール方法をひととおり実行するには約 5 分間かかります。明確で簡潔なインストール手順の例を次に示します。テストの仕様に基づいて手順を調整してください。</span><span class="sxs-lookup"><span data-stu-id="a8068-p124">On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.</span></span>

<span data-ttu-id="a8068-194">**次の手順を使用して、PowerPoint 2016 の (ここにアドイン名を挿入する) アドインをインストールしてください。**</span><span class="sxs-lookup"><span data-stu-id="a8068-194">**Please install the (insert your add-in name here) add-in for PowerPoint 2016, using the following instructions:**</span></span> 

1. <span data-ttu-id="a8068-195">Microsoft PowerPoint 2016 を開きます。</span><span class="sxs-lookup"><span data-stu-id="a8068-195">Open Microsoft PowerPoint 2016.</span></span>
2. <span data-ttu-id="a8068-196">**[新しいプレゼンテーション]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a8068-196">Select **Blank Presentation.**</span></span>
3. <span data-ttu-id="a8068-197">**[挿入] > [個人用アドイン]** に進みます。</span><span class="sxs-lookup"><span data-stu-id="a8068-197">Go to **Insert > My Add-ins.**</span></span>
5. <span data-ttu-id="a8068-198">ポップアップ ウィンドウで、**[ストア]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="a8068-198">In the popup window, choose **Store.**</span></span>
6. <span data-ttu-id="a8068-199">検索ボックスに (アドイン名) と入力します。</span><span class="sxs-lookup"><span data-stu-id="a8068-199">Type (Add-in name) in the search box.</span></span>
7. <span data-ttu-id="a8068-200">(アドイン名) を選択します。</span><span class="sxs-lookup"><span data-stu-id="a8068-200">Choose (Add-in name).</span></span>
8. <span data-ttu-id="a8068-201">少し時間を取って [ストア] ページを参照し、アドインに精通します。</span><span class="sxs-lookup"><span data-stu-id="a8068-201">Take a moment to look at the Store page to familiarize yourself with the add-in.</span></span>
9. <span data-ttu-id="a8068-202">**[追加]** を選択して、アドインをインストールします。</span><span class="sxs-lookup"><span data-stu-id="a8068-202">Choose **Add** to install the add-in.</span></span>

<span data-ttu-id="a8068-p125">いずれの相互作用や表示の忠実性のレベルでもプロトタイプをテストできます。リンクや相互作用がさらに複雑な場合は、[InVision](https://www.invisionapp.com) のようなプロトタイプ作成ツールを検討してください。静的な画面だけテストする場合は、オンラインで画像をホスティングして対応する URL を参加者に送信したり、オンラインの PowerPoint プレゼンテーションへのリンクを提供したりできます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p125">You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation.</span></span> 

## <a name="7-run-a-pilot-test"></a><span data-ttu-id="a8068-206">7.パイロット テストを実行する</span><span class="sxs-lookup"><span data-stu-id="a8068-206">7. Run a pilot test</span></span>

<span data-ttu-id="a8068-p126">プロトタイプや、タスクと質問の一覧を正しく作成するには、テクニックを要する場合があります。ユーザーがタスクのために混乱したり、プロトタイプで途方にくれる可能性があります。1 人から 3 人のユーザーを用いてパイロット テストを実行し、テストの形式に関する避けられない問題を解決する必要があります。こうすると確実に、質問を明確にし、プロトタイプを正しく設定し、探している種類のデータをキャプチャすることができます。</span><span class="sxs-lookup"><span data-stu-id="a8068-p126">It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.</span></span>

## <a name="8-run-the-test"></a><span data-ttu-id="a8068-211">8.テストを実行する</span><span class="sxs-lookup"><span data-stu-id="a8068-211">8. Run the test</span></span>

<span data-ttu-id="a8068-p127">テストを指示した後に、参加者がそのテストを完了すると、電子メール通知を受け取ります。特定の参加者のグループがターゲットになっている場合を除き、通常テストは数時間以内に完了します。</span><span class="sxs-lookup"><span data-stu-id="a8068-p127">After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.</span></span>

## <a name="9-analyze-results"></a><span data-ttu-id="a8068-214">9.結果を分析する</span><span class="sxs-lookup"><span data-stu-id="a8068-214">9. Analyze results</span></span>

<span data-ttu-id="a8068-p128">この時点で、収集したデータが意味のあるものになります。テストのビデオを見ている間に、ユーザーが直面した問題や成功についてメモを記録します。すべての結果を見終えるまで、データの意味を解釈しようとしないでください。</span><span class="sxs-lookup"><span data-stu-id="a8068-p128">This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.</span></span> 

<span data-ttu-id="a8068-p129">ユーザビリティの問題を抱える参加者が 1 人では、デザインに変更を加える正当な理由としては十分ではありません。複数の参加者が同じ問題に直面している場合は、一般集団内の他のユーザーもその問題に直面するであろうことを示しています。</span><span class="sxs-lookup"><span data-stu-id="a8068-p129">A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.</span></span>

<span data-ttu-id="a8068-p130">一般に、データを使用して結論を出す方法については注意が必要です。特定のストーリーにデータを合わせようとする誤ちに注意してください。実際にデータが証明していること、反証していること、単に洞察を提供できないことについて公正に吟味してください。先入観を持たないでください。ユーザーの動作はデザイナーの期待に反することがよくあります。</span><span class="sxs-lookup"><span data-stu-id="a8068-p130">In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.</span></span>
 

## <a name="see-also"></a><span data-ttu-id="a8068-223">関連項目</span><span class="sxs-lookup"><span data-stu-id="a8068-223">See also</span></span>
 
 - [<span data-ttu-id="a8068-224">ユーザビリティ テストを実施する方法</span><span class="sxs-lookup"><span data-stu-id="a8068-224">How to Conduct Usability Testing</span></span>](http://whatpixel.com/howto-conduct-usability-testing/)  
 - [<span data-ttu-id="a8068-225">ベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="a8068-225">Best Practices</span></span>](http://help.usertesting.com/customer/en/portal/articles/1680726-best-practices)  
 - [<span data-ttu-id="a8068-226">偏りを最小限に抑える</span><span class="sxs-lookup"><span data-stu-id="a8068-226">Minimizing Bias</span></span>](http://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
