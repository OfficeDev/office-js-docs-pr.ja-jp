# <a name="contribute-to-this-documentation"></a>このドキュメントに投稿する

このドキュメントをご感心をお寄せいただき、ありがとうございます。

* [投稿する方法](#ways-to-contribute)
* [GitHub を使用して投稿する](#contribute-using-github)
* [Git を使用して投稿する](#contribute-using-git)
* [Markdown を使用してトピックを書式設定する方法](#how-to-use-markdown-to-format-your-topic)
* [FAQ](#faq)
* [その他のリソース](#more-resources)

## <a name="ways-to-contribute"></a>投稿する方法

このドキュメントに投稿するいくつかの方法を下に示します。

* 記事に小さな変更を加える方法については、「[GitHub を使用して投稿する](#contribute-using-github)」を参照してください。
* 大きな変更やコードが関係する変更を加える方法については、「[Git を使用して投稿する](#contribute-using-git)」を参照してください。
* GitHub の [問題] を使用してドキュメントの不具合を報告します。
* 「[Office Developer Platform UserVoice](http://officespdev.uservoice.com)」サイトで新しいドキュメントを要求します。

## <a name="contribute-using-github"></a>GitHub を使用して投稿する

リポジトリをデスクトップに複製せずにこのドキュメントに投稿するには、GitHub を使用します。これは、このリポジトリでプル リクエストを作成する最も簡単な方法です。コードの変更に関係しない小さな変更を加えるには、この方法を使用します。 

**注**: この方法では、一度に 1 つの記事に投稿できます。

### <a name="to-contribute-using-github"></a>GitHub を使用して投稿するには

1. 投稿する記事を GitHub で検索します。
2. GitHub で記事が表示されたら、GitHub にサインインします (無料アカウントを取得するには、「[Join GitHub](https://github.com/join)」 (GitHub に参加) にアクセスします)。
3. **鉛筆アイコン** (このプロジェクトのフォークでファイルを編集します) を選択し、**[<>ファイルの編集]** ウィンドウで変更を加えます。 
4. 一番下までスクロールし、説明を入力します。
5. [**ファイル変更の提案**] > [**プル リクエストの作成**] を選択します。

これでプル リクエストを正常に提出できました。プル リクエストは、通常 10 営業日以内に審査されます。 


## <a name="contribute-using-git"></a>Git を使用して投稿する

次のような実質的な変更を投稿するには、Git を使用します。

* コードの投稿。
* 意味に影響する変更の投稿。
* テキストの大規模な変更の投稿。
* 新しいトピックの追加。

### <a name="to-contribute-using-git"></a>Git を使用して投稿するには

1. GitHub アカウントを持っていない場合は、[GitHub](https://github.com/join) でセットアップします。 
2. アカウントを取得したら、ご利用のコンピューターに Git をインストールします。 「[Set up Git]」 (Git の設定) チュートリアルの手順を実行します。
3. Git を使用してプル要求を送信するには、「[GitHub、Git、およびこのリポジトリを使用する](#use-github-git-and-this-repository)」の手順を実行します。
4. 次の場合は、投稿者のライセンス同意書に署名するように求められます。

    * Microsoft Open Technologies グループのメンバーである
    * Microsoft の従業員でない投稿者である

コミュニティ メンバーは、プロジェクトへの大規模な投稿を行う前に投稿者のライセンス同意書 (CLA) に署名する必要があります。このドキュメントに記入して送信する必要があるのは 1 回だけです。注意深く確認してください。雇用主がこのドキュメントに署名することが要求される場合もあります。

CLA への署名により、メイン リポジトリにコミットする権限が付与されるわけではありませんが、Office Developer および Office Developer Content Publishing チームからお客様の投稿への確認と承認を受けることができるようになります。 送信内容には自身の名義が入ります。

通常、プル要求は 10 営業日以内に審査されます。

## <a name="use-github-git-and-this-repository"></a>GitHub、Git、およびこのリポジトリを使用する

**注:** このセクションの情報のほとんどは、「[GitHub Help]」 (GitHub ヘルプ) の記事にあります。  Git と GitHub のことをよく知っている場合は、「**コンテンツを投稿して編集する**」のセクションまでスキップして、このリポジトリのコード/コンテンツ フローの詳細を参照してください。

### <a name="to-set-up-your-fork-of-the-repository"></a>リポジトリのフォークを設定するには

1.  このプロジェクトに投稿できるように、GitHub のアカウントをセットアップします。まだ行っていない場合は、今すぐ [GitHub](https://github.com/join) にアクセスしてセットアップします。
2.  ご利用のコンピューターに Git をインストールします。 「[Set up Git]」 (Git の設定) チュートリアルの手順を実行します。
3.  このリポジトリの独自のフォークを作成します。 これを行うには、ページの上部にある **[フォーク]** ボタンを選択します。
4.  フォークをコンピューターにコピーします。これを行うには、Git Bash を開きます。コマンド プロンプトで、次のように入力します。

        git clone https://github.com/<your user name>/<repo name>.git

    次に、以下のコマンドを入力してルート リポジトリへの参照を作成します。

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

おめでとうございます。リポジトリをセットアップできました。今後、同じ手順をもう一度繰り返す必要はありません。

### <a name="contribute-and-edit-content"></a>コンテンツを投稿して編集する

投稿プロセスをできるだけシームレスにするため、以下の手順に従ってください。

#### <a name="to-contribute-and-edit-content"></a>コンテンツを投稿して編集するには

1. 新しい分岐を作成します。
2. 新しい内容を追加するか、既存の内容を編集します。
3. メイン リポジトリにプル リクエストを提出します。
4. 分岐を削除します。

**重要** 作業フローを効率化してマージによる競合の可能性を減らすため、各分岐を単一の概念 / 記事に限定してください。新しい分岐に適した内容には、次のものが含まれます。

* 新しい記事。
* スペルと文法の編集。
* 大規模な記事セット全体への単一の書式設定変更の適用 (たとえば、新しい著作権フッターの適用)。

#### <a name="to-create-a-new-branch"></a>新しい分岐を作成するには

1.  Git Bash を開きます。
2.  Git Bash のコマンド プロンプトで、`git pull upstream master:<new branch name>` と入力します。これにより、最新の OfficeDev マスター分岐からコピーされた新しい分岐がローカルに作成されます。
3.  Git Bash のコマンド プロンプトで、`git push origin <new branch name>` と入力します。これにより、新しい分岐に関するアラートが GitHub に送信されます。これで、GitHub 上のリポジトリのフォーク内に新しい分岐が表示されるはずです。
4.  新しい分岐に切り替えるには、Git Bash コマンド プロンプトで `git checkout <new branch name>` と入力します。

#### <a name="add-new-content-or-edit-existing-content"></a>新しい内容を追加するか既存の内容を編集する

ファイル エクスプローラーを使用して、コンピューター上のリポジトリに移動します。リポジトリ ファイルは、`C:\Users\<yourusername>\<repo name>` にあります。

ファイルを編集するには、好みのエディターで開いて変更します。新しいファイルを作成するには、好みのエディターを使用して、リポジトリのローカル コピー内の適切な場所に新しいファイルを保存します。作業中は、頻繁に作業内容を保存してください。

`C:\Users\<yourusername>\<repo name>` 内のファイルは、ローカル リポジトリ内で作成した新しい分岐の作業コピーです。このフォルダーで行った変更は、変更をコミットするまでローカル リポジトリに影響しません。ローカル リポジトリに変更をコミットするには、GitBash で次のコマンドを入力します。

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

`add` コマンドにより、変更はリポジトリへのコミットの準備としてステージング領域に追加されます。`add` コマンドの後のピリオドは、サブフォルダーを再帰的にチェックして、追加または変更したすべてのファイルをステージングすることを指定します。(すべての変更をコミットするのでない場合は、特定のファイルを追加できます。コミットを元に戻すこともできます。ヘルプを表示するには、「`git add -help`」または「`git status`」と入力してください。)

`commit` コマンドにより、ステージングされた変更がリポジトリに適用されます。スイッチ `-m` は、コミット コメントをコマンドラインで提供することを意味します。-v および -a スイッチは省略できます。-v スイッチはコマンドからの詳細 (verbose) 出力用で、-a スイッチは add コマンドですでに行ったことを行います。

作業の途中で複数回コミットするか、完了時に 1 回コミットすることができます。

#### <a name="submit-a-pull-request-to-the-main-repository"></a>メイン リポジトリにプル リクエストを送信する

作業が完了し、メイン リポジトリにマージする準備ができたら、以下の手順を実行します。

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>メイン リポジトリにプル リクエストを送信するには

1.  Git Bash のコマンド プロンプトで、`git push origin <new branch name>` と入力します。ローカル リポジトリにおいて、`origin` はローカル リポジトリの複製元である GitHub リポジトリを指します。このコマンドにより、前の手順で行ったすべてのコミットを含む新しい分岐の現在の状態が GitHub フォークにプッシュされます。
2.  GitHub サイト上のフォーク内で、新しい分岐まで移動します。
3.  ページの上部にある [**プル リクエスト**] ボタンを選択します。
4.  Base 分岐が `OfficeDev/<repo name>@master` で、Head 分岐が `<your username>/<repo name>@<branch name>` であることを確認します。
5.  [**コミット範囲の更新**] ボタンを選択します。
6.  プル リクエストにタイトルを追加し、作成しているすべての変更についての説明を入力します。
7.  プル リクエストを提出します。

サイト管理者の 1 人がプル要求を処理します。プル要求が OfficeDev/<repo name> サイトの [問題] に表示されます。プル要求が承諾されると、問題は解決されます。

#### <a name="create-a-new-branch-after-merge"></a>マージの後に新しい分岐を作成する

分岐が正常にマージされた (つまり、プル リクエストが承諾された) 後は、ローカル分岐で作業を継続しないでください。別のプル リクエストを提出する場合にマージの競合が発生する可能性があります。別の更新を行うには、正常にマージされたアップストリーム分岐から新しいローカル分岐を作成した後、最初のローカル分岐を削除します。

たとえば、ローカル分岐 X が正常に OfficeDev/microsoft-graph-docs マスター分岐にマージされた後、マージされた内容に追加の更新を行う場合を考えます。OfficeDev/microsoft-graph-docs マスター分岐から新しいローカル分岐 X2 を作成します。これを行うには、GitBash を開き、次のコマンドを実行します。

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

これで、分岐 X で提出した作業のローカル コピーを (新しいローカル分岐内に) 作成できました。X2 ブランチには他のライターがマージしたすべての作業も含まれるため、自身の作業が他のライターの作業 (たとえば、共有画像) に依存している場合はその作業が新しい分岐で使用可能になります。以前の作業 (および他のライターの作業) が分岐にあることを確認するには、新しい分岐をチェックアウトして...

    git checkout X2

...内容を確認します。(`checkout` コマンドは、`C:\Users\<yourusername>\microsoft-graph-docs` 内のファイルを X2 分岐の現在の状態に更新します。新しい分岐をチェックアウトした後は、内容への更新を行って通常どおりにコミットすることができます。ただし、誤って作業することを避けるため、結合済みの分岐 (X) は削除することをお勧めします (次の「**分岐を削除する**」のセクションを参照してください)。

#### <a name="delete-a-branch"></a>分岐を削除する

変更内容がメイン リポジトリにマージされたら、使用した分岐は不要になったので削除します。追加の作業は新しい分岐で行う必要があります。  

#### <a name="to-delete-a-branch"></a>分岐を削除するには

1.  Git Bash のコマンド プロンプトで、「`git checkout master`」と入力します。これにより、削除される分岐にいないことが保証されます (削除される分岐にいることは許可されません)。
2.  次に、コマンド プロンプトで `git branch -d <branch name>` と入力します。これにより、アップストリーム リポジトリに正常にマージ済みの場合にのみ、コンピューター上の分岐が削除されます。(この動作は `–D` フラグでオーバーライドすることができますが、最初にその必要があるかどうかを確かめてください)。
3.  最後に、コマンド プロンプトで `git push origin :<branch name>` (コロンの前にはスペースを 1 つ入れ、後にはスペースを入れません) と入力します。これにより、GitHub フォーク上の分岐が削除されます。  

おめでとうございます。プロジェクトに正しく投稿できました。

## <a name="how-to-use-markdown-to-format-your-topic"></a>Markdown を使用してトピックを書式設定する方法

### <a name="markdown"></a>Markdown

このリポジトリ内のすべての記事では、Markdown を使用しています。 完全な紹介 (および、すべての構文のリスト) は、「[Daring Fireball - Markdown]」にあります。
 
## <a name="faq"></a>FAQ

### <a name="how-do-i-get-a-github-account"></a>GitHub アカウントを取得する方法を教えてください。

無料の GitHub アカウントを開設するには、「[Join GitHub](https://github.com/join)」(GitHub に参加) にあるフォームに記入します。 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>投稿者のライセンス同意書はどこで入手するのでしょうか。 

プル リクエストで投稿者のライセンス同意書 (CLA) が必要な場合、CLA の署名が必要であることを述べる通知が自動的に送信されます。 

コミュニティ メンバーは、**このプロジェクトへの大規模な投稿を行う前に投稿者のライセンス同意書 (CLA) に署名する必要があります**。このドキュメントに記入して送信する必要があるのは 1 回だけです。注意深く確認してください。雇用主がこのドキュメントに署名することが要求される場合もあります。

### <a name="what-happens-with-my-contributions"></a>私が投稿した内容はどうなりますか。

プル リクエストを使用して変更を提出すると、弊社チームがその通知を受け、プル リクエストを審査します。投稿者には、プル リクエストに関する通知が GitHub から送られます。さらに情報が必要な場合、弊社チームのメンバーからも通知が送られます。プル リクエストが承認された場合、ドキュメントを更新します。弊社は、投稿された内容を、法律、スタイル、わかりやすさ、またはその他の理由で編集する権利を保持します。

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>このリポジトリの GitHub プル リクエストの承認者になることができますか。

現在、外部の投稿者がこのリポジトリ内のプル リクエストを承認することは許可されていません。

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>変更リクエストに関する応答をどのくらいの期間内に受けることができますか。

プル リクエストは、通常 10 営業日以内に審査されます。


## <a name="more-resources"></a>その他のリソース

* Markdown に関する詳細については、Markdown 作成者のサイト「[Daring Fireball]」にアクセスしてください。
* Git と GitHub の使用に関する詳細については、まず「[GitHub Help]」 (GitHub ヘルプ) を確認してください。

[GitHub Home]: http://github.com
[GitHub ヘルプ]: http://help.github.com/
[Git の設定]: https://help.github.com/articles/set-up-git/
[Daring Fireball - Markdown]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
