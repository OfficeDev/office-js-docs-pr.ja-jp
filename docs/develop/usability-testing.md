---
title: Office アドインのユーザビリティ テスト
description: 実際のユーザーでアドインの設計をテストする方法について説明します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 49a2af983615779160886961e8269e4588d0fc9e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810282"
---
# <a name="usability-testing-for-office-add-ins"></a>Office アドインのユーザビリティ テスト

A great add-in design takes user behaviors into account. Because your own preconceptions influence your design decisions, it’s important to test designs with real users to make sure that your add-ins work well for your customers.

使いやすさテストはさまざまな方法で実行できます。 多くのアドイン開発者にとって、リモートでモデレートされていないユーザビリティスタディが最も時間とコスト効率に優れています。 いくつかの一般的なテスト サービスにより、この操作が簡単になります。いくつかの例を次に示します。

- [UserTesting.com](https://www.UserTesting.com)
- [Optimalworkshop.com](https://www.Optimalworkshop.com)
- [Userzoom.com](https://www.Userzoom.com)

これらのテスト サービスは、テスト計画の作成を効率化し、参加者を探したりテストをモデレートしたりする必要をなくすのに役立ちます。

You need only five participants to uncover most usability issues in your design. Incorporate small tests regularly throughout your development cycle to ensure that your product is user-centered.

> [!NOTE]
> We recommend that you test the usability of your add-in across multiple platforms. To [publish your add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center), it must work on all [platforms that support the methods that you define](/javascript/api/requirement-sets).

## <a name="1-sign-up-for-a-testing-service"></a>1. テスト サービスにサインアップする

詳細については、「[Selecting an Online Tool for Unmoderated Remote User Testing](https://www.nngroup.com/articles/unmoderated-user-testing-tools/)」 (モデレートされていないリモート ユーザー テスト用のオンライン ツールの選択) を参照してください。

## <a name="2-develop-your-research-questions"></a>2. 調査での質問項目を設定する

Research questions define the objectives of your research and guide your test plan. Your questions will help you identify participants to recruit and the tasks they will perform. Make your research questions as specific as you can. You can also seek to answer broader questions.

研究の質問の例を次に示します。

**具体的な質問**

- ユーザーは、ランディング ページ上の "無料試用版" リンクに気が付きますか。
- ユーザーがアドインから自分のドキュメントにコンテンツを挿入する際に、ドキュメント内の挿入場所がわかるでしょうか。

**広範な質問**

- ユーザーにとってアドインの最大の懸案事項は何ですか。
- ユーザーは、コマンド バー内のアイコンの意味を、クリックする前に理解していますか。
- ユーザーは設定メニューを簡単に検索できますか。

アドインを見つけることからインストールして使用することに至るまで、ユーザー体験全体に関するデータを取得することは重要です。 アドイン ユーザー エクスペリエンスの次の側面に対処する調査の質問を検討してください。

- AppSource 内でのアドインの検索
- アドインのインストールの選択
- 最初の実行エクスペリエンス
- リボン コマンド
- アドイン UI
- アドインが Office アプリケーションのドキュメント領域と相互作用する方法
- ユーザーがコンテンツ挿入フローを制御できる程度

詳細については、「[Gathering factual responses vs. subjective data](https://help.usertesting.com/hc/articles/115003378572-Writing-effective-questions)」 (実際の反応と主観的データを収集する) を参照してください。

## <a name="3-identify-participants-to-target"></a>3. ターゲットとする参加者を特定する

Remote testing services can give you control over many characteristics of your test participants. Think carefully about what kinds of users you want to target. In your early stages of data collection, it might be better to recruit a wide variety of participants to identify more obvious usability issues. Later, you might choose to target groups like advanced Office users, particular occupations, or specific age ranges.

## <a name="4-create-the-participant-screener"></a>4.参加者のスクリーナーを作成する

The screener is the set of questions and requirements you will present to prospective test participants to screen them for your test. Keep in mind that participants for services like UserTesting.com have a financial interest in qualifying for your test. It's a good idea to include trick questions in your screener if you want to  exclude certain users from the test. 

たとえば、GitHub に精通している参加者を見つけようとしている場合に、身元を偽っている可能性があるユーザーを除外するには、考えられる回答の一覧に偽の回答を組み込みます。

**次のソース コード リポジトリのうちどれに精通していますか。**  
 a. SourceShelf  [*Reject*]  
 b. CodeContainer  [*Reject*]  
 c. GitHub  [*Must select*]  
 d. BitBucket  [*May select*]  
 e. CloudForge  [*May select*]  

アドインの Live ビルドのテストを計画している場合は、次の質問により、このテストを行えるユーザーを選別できます。

**このテストでは、Microsoft PowerPoint の最新バージョンを所有している必要があります。最新バージョンの PowerPoint を使用していますか。**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  
 c. I don’t know [*Reject*]  

**このテストでは、PowerPoint の無料のアドインをインストールし、これを使用するために無料のアカウントを作成する必要があります。アドインをインストールし、無料のアカウントを作成する予定がありますか。**  
 a. Yes [*Must select*]  
 b. No [*Reject*]  

詳細については、「[Screener Questions Best Practices](https://help.usertesting.com/hc/articles/115003370731-Screener-question-best-practices)」 (スクリーナーの質問のベスト プラクティス) を参照してください。

## <a name="5-create-tasks-and-questions-for-participants"></a>5. 参加者のタスクと質問を作成する

Try to prioritize what you want tested so that you can limit the number of tasks and questions for the participant. Some services pay participants only for a set amount of time, so you want to make sure not to go over.

Try to observe participant behaviors instead of asking about them, whenever possible. If you need to ask about behaviors, ask about what participants have done in the past, rather than what they would expect to do in a situation. This tends to give more reliable results.

The main challenge in unmoderated testing is making sure your participants understand your tasks and scenarios. Your directions should be *clear and concise*. Inevitably, if there is potential for confusion, someone will be confused.

Don't assume that your user will be on the screen they’re supposed to be on at any given point during the test. Consider telling them what screen they need to be on to start the next task.

詳細については、「[Writing Great Tasks](https://help.usertesting.com/hc/articles/115003371651-Writing-great-tasks)」 (優れたタスクの作成) を参照してください。

## <a name="6-create-a-prototype-to-match-the-tasks-and-questions"></a>6. タスクや質問に対応するプロトタイプを作成する

Live アドインをテストするか、プロトタイプをテストすることができます。 Live アドインをテストする場合は、Office の最新バージョンがあり、アドインのインストールを予定しており、アカウントのサインアップを予定している (ログオン資格情報を提供しない場合) 参加者を選別する必要があることに注意してください。後で、アドインを正常にインストールしたかどうかを確認する必要があります。

On average, it takes about 5 minutes to walk users through how to install an add-in. The following is an example of clear, concise installation steps. Adjust the steps based on the specifics of your test.

**次の手順に従って、PowerPoint 用アドインをインストールしてください (ここにアドイン名を挿入してください)。**

1. Microsoft PowerPoint を開きます。
1. **[新しいプレゼンテーション]** を選択します。
1. [ **アドインの挿入]** > **に移動します**。
1. ポップアップ ウィンドウで、[ストア] を選択 **します**。
1. 検索ボックスに (アドイン名) と入力します。
1. (アドイン名) を選択します。
1. 少し時間を取って [ストア] ページを参照し、アドインに精通します。
1. **[追加]** を選択して、アドインをインストールします。

You can test a prototype at any level of interaction and visual fidelity. For more complex linking and interactivity, consider a prototyping tool like [InVision](https://www.invisionapp.com). If you just want to test static screens, you can host images online and send participants the corresponding URL, or give them a link to an online PowerPoint presentation. 

## <a name="7-run-a-pilot-test"></a>7.パイロット テストを実行する

It can be tricky to get the prototype and your task/question list right. Users might be confused by tasks, or might get lost in your prototype. You should run a pilot test with 1-3 users to work out the inevitable issues with the test format. This will help to ensure that your questions are clear, that the prototype is set up correctly, and that you’re capturing the type of data you’re looking for.

## <a name="8-run-the-test"></a>8.テストを実行する

After you order your test, you will get email notifications when participants complete it. Unless you’ve targeted a specific group of participants, the tests are usually completed within a few hours.

## <a name="9-analyze-results"></a>9.結果を分析する

This is the part where you try to make sense of the data you’ve collected. While watching the test videos, record notes about problems and successes the user has. Avoid trying to interpret the meaning of the data until you have viewed all the results.

A single participant having a usability issue is not enough to warrant making a change to the design. Two or more participants encountering the same issue suggests that other users in the general population will also encounter that issue.

In general, be careful about how you use your data to draw conclusions. Don’t fall into the trap of trying to make the data fit a certain narrative; be honest about what the data actually proves, disproves, or simply fails to provide any insight about. Keep an open mind; user behavior frequently defies designer’s expectations.

## <a name="see-also"></a>関連項目

- [ユーザビリティ テストを実施する方法](https://whatpixel.com/howto-conduct-usability-testing/)  
- [UserTesting のベスト プラクティス](https://help.usertesting.com/hc/articles/115003370231-Best-practices-for-UserTesting)  
- [偏りを最小限に抑える](https://downloads.usertesting.com/white_papers/TipSheet_MinimizingBias.pdf)  
