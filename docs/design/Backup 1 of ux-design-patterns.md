---
title: Office アドインの UX 設計パターン テンプレート
description: ''
ms.date: 01/23/2018
---



# <a name="ux-design-pattern-templates-for-office-add-ins"></a>Office アドインの UX 設計パターン テンプレート 

[Office アドインの UX 設計パターン プロジェクト](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "Office アドインの UX 設計パターン プロジェクト")には、アドインの UX を作成できる HTML、JavaScript、および CSS ファイルが含まれています。   

UX 設計パターン プロジェクトは、次に使用できます。

* よくある顧客のシナリオにソリューションとして適用する。
* 設計のベスト プラクティスとして適用する。
* [Office UI Fabric](https://dev.office.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。
* Office の既定の UI に視覚的に溶け込むアドインをビルドする。  

## <a name="using-the-ux-design-patterns"></a>UX 設計パターンの使用

[UX 設計パターン仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns)を独自の Office アドインを設計する際のガイドとして使用することも、[ソース コード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates)をプロジェクトに直接追加することもできます。

この仕様を使用して、独自のアドイン UI に模擬表示をビルドするには、次を実行します。

1. 設計リソース ファイルをダウンロードし、独自 UI の設計を開始します。
    * [Office アドイン UX 設計コンポーネント](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_components.ai) (Adobe Illustrator ファイル)
    * [Office アドイン UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_patterns.ai) (Adobe Illustrator ファイル) または
    * [Office アドイン UX 設計プロトタイプ](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/addin_ux_design_prototype.xd) (Adobe Experience Design ファイル)
2. ガイダンスについては、次の記事を参照してください。
    * [UX 設計パターン](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/README.md)
    * [Office アドイン設計](add-in-design.md)のベスト プラクティス
    * 
  [Office UI Fabric ツールキット](https://developer.microsoft.com/en-us/fabric#/resources)

ソース コードを追加するには、次のようにします。

1. [Office アドインの UX 設計パターン プロジェクト リポジトリ](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code "Office アドインの UX 設計パターン プロジェクト")を複製します。 
2. [資産フォルダー](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets)と、アドイン プロジェクトに対して選ぶ個々のパターンのコード フォルダーをコピーします。  
3. 個々のパターンをアドインに組み込みます。たとえば次のようにします。
    - マニフェスト内で、ソースの場所またはアドイン コマンドの URL を編集します。
    - 他のページのテンプレートとして、UX 設計パターンを使用します。
    - UX 設計パターンとの間にリンクを設定します。

> [!NOTE]
> 一部の UX パターン仕様は、ソース コードに一致しません。 すべての資産を標準化するよう努力しています。 また、一部の仕様がアーカイブとして提供されることに注意してください。 プラットフォームにとっての価値について、これらのアーカイブされた仕様を評価しています。 各パターンは、固有のテンプレートと相互作用のパターンを表現することを目的としています。 パターンは互いに重複しないようにし、Office Fabric UI コンポーネントとは明確に区別する必要があります。

## <a name="types-of-ux-design-patterns"></a>UX 設計パターンの種類
### <a name="generic-pages"></a>汎用ページ

汎用ページ テンプレートは、アドインの任意のページに適用でき、特定の目的を持ちません。特定の目的を持つページの例は、すべての初回実行時のパターンです。使用可能な汎用ページの一覧は、以下のとおりです。

* **ランディング ページ**: 初回実行時またはサインイン時にユーザーに対して表示される標準のアドイン ページです。 
    * アドインで [Office デザイン言語](add-in-design-language.md)を採用するためのガイドラインについて説明します。
    * [ランディング ページのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page)
* **ブランド バーのブランド イメージ**: フッターに自分のブランドを表す画像を付加されたランディング ページです。 
    * [ブランド バーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/brand-bar.md)
    * [ブランド バーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar)

<table>
 <tr><th>ランディング</th><th>ブランド バー</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/landing-page"><img src="../images/landing-pages.png" alt="landing page" style="width: 264px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/generic/brand-bar"><img src="../images/word-brand-bar.png" alt="brand bar" style="width: 264px;"/></A></td></tr>
 </table>
 
### <a name="first-run-experience"></a>初回実行時エクスペリエンス

初回実行時エクスペリエンスとは、ユーザーが最初にアドインを開いたときのエクスペリエンスです。初回実行時の設計パターン テンプレートは、以下のとおりです。 

* **開始手順**: アドインの使用を開始する手順の、順序付きリストをユーザーに提供します。 
    * [開始手順の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_stepsToStart.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときには、[初回実行時の価値の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)を参照してください)。  
    * [開始手順のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step)
* **価値**: アドインの価値提供を明確にします。
    * [価値の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/value-placemat.md)
    * [価値のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat)
* **ビデオ**: アドインの使用を開始する前に、ユーザーにビデオを表示します。
    * [ビデオの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/video-placemat.md)
    * [ビデオのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat)
* **チュートリアル**: アドインの使用を開始する前に、ユーザーに一連の機能または情報を体験させます。
    * [カルーセルの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/carousel.md) (この UX 設計パターンの名前は「カルーセル」に変更されています。 以前の仕様では、「ページング パネル」と呼ばれていました。 コード アセットでは、「初回実行時チュートリアル」と呼びます。 
    * [チュートリアルのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough)


  [AppSource](https://docs.microsoft.com/ja-jp/office/dev/store/use-the-seller-dashboard-to-submit-to-the-office-store) には、アドインの試用版を管理するシステムが存在しますが、アドインの試用版エクスペリエンスの UI を自分で管理する場合は、次のパターンを使用します。

* **試用**: アドインの試用版で開始する方法をユーザーに示します。
    * [試用の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialVersion.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [試用のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat)
* **試用版の機能**: ユーザーが試用を考えている機能がアドインの試用版では使用できないことをユーザーに示します。または、無料のアドインにサブスクリプションを必要とする機能も含まれる場合は、このパターンの使用を検討します。このパターンでは、試用期間の終了後にダウングレードしたエクスペリエンスを提供することもできます。
    * [試用版の機能の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/fre_trialFeature.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [試用版の機能のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature)

> [!IMPORTANT]
> 自分で試用版を管理し、試用版の管理に AppSource を使用しない場合は、必ず販売者ダッシュボードのテスト用メモに、**[追加購入が必要になる場合があります]** タグを組み込んでください。

ユーザーに初回実行時エクスペリエンスを 1 回示すか、何回も示すかを検討することがシナリオにとって重要かどうかを検討します。たとえば、ユーザーがアドインを定期的に使用する場合は、使用方法を忘れる可能性があるため、初回実行エクスペリエンスを複数回確認できた方がよいでしょう。 

 <table>
 <tr><th>開始手順</th><th>値</th><th>ビデオ</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/instruction-step"><img src="../images/instruction-steps.png" alt="instruction steps" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/value-placemat"><img src="../images/value-placemats.png" alt="value placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/video-placemat"><img src="../images/video-placemats.png" alt="video placemat" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>チュートリアルの最初のページ</th><th>試用</th><th>試用版の機能</th></tr>
 <tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/walkthrough"><img src="../images/walkthrough01.png" alt="walkthrough 1" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat"><img src="../images/trial-placemats.png" alt="trial placemat" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/first-run/trial-placemat-feature"><img src="../images/trial-placemats-feature.png" alt="trial placemat feature" style="width: 250px;"/></A></td></tr>
 </table> 

### <a name="navigation"></a>ナビゲーション

ユーザーは、アドインの別のページ間を移動する必要があります。次のナビゲーション テンプレートには、アドインのページおよびコマンドの整理に使用できるさまざまなオプションがあります。

* **[戻る] ボタンと [次のページ]**: [戻る] ボタンと [次のページ] がある作業ウィンドウを表示します。ユーザーが順序のある一連の手順に従えるようにするには、このパターンを使用します。
    * [[戻る] ボタンと [次のページ] の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/back-button.md)
    * [[戻る] ボタンと [次のページ] のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button) 
* **ナビゲーション**: 作業ウィンドウにページのメニュー項目がある、一般にハンバーガー メニューと呼ばれるメニューを表示します。 
    * [ナビゲーションの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/contextual-menu.md)
    * [ナビゲーションのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation) 
* **コマンド付きナビゲーション**: 作業ウィンドウにコマンド (またはアクション) ボタンが付いたハンバーガー メニューを表示します。ナビゲーションにコマンド オプションも必要な場合は、このパターンを使用します。 
    * [コマンド付きナビゲーションの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/command-bar.md)
    * [コマンド付きナビゲーションのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands)
* **ピボット**: 作業ウィンドウにピボットでのナビゲーションを表示します。ピボット ナビゲーションを使用すると、ユーザーは異なるコンテンツ間を移動できます。
    * [ピボットの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/pivot.md)
    * [ピボットのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot)
* **タブ バー**: テキストとアイコンが縦に並んだボタンが使用されたナビゲーションを表示します。タブ バーを使用すると、短くてわかりやすいタイトルのタブが使用されたナビゲーションを表示できます。
    * [タブ バーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/tab-bar.md)
    * [タブ バーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar) 

<table>
<tr><th>[戻る] ボタン</th><th>ナビゲーション</th><th>コマンド付きナビゲーション</th></tr>
<tr>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/back-button">
        <img src="../images/back-button.png" alt="back button" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation">
        <img src="../images/navigation.png" alt="navigation" style="width: 250px;"/></A>
    </td>
    <td>
        <A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/navigation-commands">
        <img src="../images/navigation-commands.png" alt="navigation with commands" style="width: 250px;"/></A>
    </td>
</tr>
 </table>

<table>
<tr><th>ピボット</th><th>タブ バー</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/pivot">
<img src="../images/pivot.png" alt="pivot navigation" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/navigation/tab-bar">
<img src="../images/tab-bar.png" alt="tab bar" style="width: 250px;"/></A></td>
</tr>
 </table>

### <a name="notifications"></a>通知

アドインでは、ユーザーにさまざまな方法で、エラーなどのイベントや進捗状況を通知できます。使用可能な通知テンプレートは次のとおりです。 

* **埋め込みダイアログ ボックス**: ボタンまたはその他のコントロールを使用して、情報や必要に応じて対話型エクスペリエンスを提供するダイアログ ボックスを作業ウィンドウ内に表示します。いずれか 1 つを使用して、ユーザーにアクションの確認を促すことを検討します。ユーザー エクスペリエンスを作業ウィンドウに維持したい場合は、埋め込みダイアログ パターンを使用します。
    * [埋め込みダイアログ ボックスの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/embedded-dialog.md)
    * [埋め込みダイアログ ボックスのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog)
* **インライン メッセージ** - エラー、成功、情報を示します。メッセージは作業ウィンドウ内の指定した場所に表示できます。たとえば、ユーザーがテキスト ボックスに不適切な書式の電子メール アドレスを入力すると、テキスト ボックスの真下にエラー メッセージが表示されます。 
    * [インライン メッセージの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_inlineMessage.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [インライン メッセージのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message)
* **メッセージ バナー**: 情報や必要に応じてアクションの単純な呼び出しを、単一の行に折りたたんだり、複数の行に展開したり、非表示にできるバナーで提供します。メッセージ バナーは、サービスの更新またはアドインを開始するときに役に立つヒントの報告に使用します。 
    * [メッセージ バナーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/message_bar.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [メッセージ バナーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner)
* **進行状況バー**: ユーザーが以降のアクションを行う前に完了する必要がある、実行時間の長い、同期プロセス (構成タスクなど) の進行状況を示します。これは別のスポット ページであり、アドインのブランド化を強化します。進行状況バーは、アドインに戻るまでの定期的な評価をプロセスが送信する際に使用します。
    * [進行状況バーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/progress-indicator.md)
    * [進行状況バーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar)
* **スピナー**: 実行時間の長い、同期プロセスが進行中だが、どの程度進んでいるのかは提示されていないことを示します。これは別のスポット ページであり、アドインのブランド化を強化します。スピナーは、プロセスがどの程度進んでいるかをアドインが確実に知ることができない場合に使用します。 
    * [スピナーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/spinner.md)
    * [スピナーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner)
* **トースト**: 数秒で消える簡単なメッセージを提供します。ユーザーがメッセージに気づかない場合があるため、トーストは重要ではない情報にのみ使用します。これは、電子メールの受信など、リモート システムでユーザーにイベントを通知する場合に優れた選択になります。
    * [トーストの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/toast.md)
    * [トーストのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast)

 <table>
 <tr><th>埋め込みのダイアログ</th><th>インライン メッセージ</th><th>メッセージ バナー</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/embedded-dialog"><img src="../images/embedded-dialogs.png" alt="embedded dialog" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/inline-message"><img src="../images/inline-messages.png" alt="inline message" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/message-banner"><img src="../images/message-banners.png" alt="message banner" style="width: 250px;"/></A></td></tr>
 </table>

 <table>
 <tr><th>進行状況バー</th><th>スピナー</th><th>トースト</th></tr>
 <tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/progress-bar"><img src="../images/progress-bars.png" alt="progress bar" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/spinner"><img src="../images/logo-spinner.png" alt="spinner" style="width: 250px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/notifications/toast"><img src="../images/toast-header.png" alt="toast" style="width: 250px;"/></A></td></tr>
 </table>
 


### <a name="general-components"></a>一般的なコンポーネント

以下に、さまざまなシナリオで使用できるアドイン用の一般的なコンポーネントを示します。  

#### <a name="client-dialog-boxes"></a>クライアント ダイアログ ボックス

クライアント ダイアログ ボックスは、ユーザーが作業ウィンドウ外でアドインを操作できる別の方法を提供します。使用可能なダイアログ ボックス テンプレートは次のとおりです。

* **Typeramp ダイアログ ボックス**: テキスト コンテンツを含むダイアログ ボックスを表示します。Typeramp ダイアログを使用すると、ユーザーに詳細な情報を表示できます。 
    * [Office アドインのダイアログ ボックス](dialog-boxes.md)の設計について説明します。[Office アドインの文字体裁](add-in-design-language.md#typography)に関するガイドラインにも従ってください。
    * [Typeramp ダイアログ ボックスのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp)
* **警告ダイアログ ボックス**: ユーザーへのエラーや通知などの重要な情報を含む警告ボックスを表示します。  
    * [警告ダイアログ ボックスの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_alert.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [警告ダイアログ ボックスのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert)
* **ナビゲーション ダイアログ ボックス**: ナビゲーションを含むダイアログ ボックスを表示します。ナビゲーション ダイアログ ボックスを使用すると、ユーザーは異なるコンテンツ間を移動できます。 
    * [Office アドインのダイアログ ボックス](dialog-boxes.md)の設計について説明します。また、Office UI Fabric の [Office アドインのピボット コンポーネント](pivot.md)の使用について説明します。
    * [ナビゲーション ダイアログ ボックスのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation)

<table>
 <tr><th>Typeramp ダイアログ</th><th>警告ダイアログ</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/typeramp"><img src="../images/typeramp-dialog.png" alt="typeramp dialog" width="400"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/alert"><img src="../images/alert-dialog.png" alt="alert dialog" width="400"/></A></td>
</tr></tr>
 </table>
 
 <table>
 <tr><th>ナビゲーション ダイアログ</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/dialog/navigation"><img src="../images/navigation-dialog.png" alt="navigation dialog" width="450"/></A></td></tr>
</tr>
 </table>


#### <a name="feedback-and-ratings"></a>フィードバックおよび評価

アドインの可視性や導入を向上するには、AppSource でユーザーがアドインを評価およびレビューできる機能を提供することをお勧めします。このパターンには、アドインからフィードバックおよび評価を提示できる次の 2 つの方法があります。

- ユーザー側からのフィードバック: ([**フィードバックの送信**] リンクなどの) ナビゲーション メニューまたはフッターのアイコンを使用し、ユーザーがフィードバックの送信を選択します。
- システム側からのフィードバック: アドインが 3 回実行された後、メッセージ バナーを介してフィードバックを行うようユーザーが促されます。

いずれの方法でも、AppSource のアドイン用のページを含むダイアログ ボックスが開きます。

* [フィードバックおよび評価の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/notification_feedback.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
* [フィードバックおよび評価のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store)

> [!IMPORTANT]
> このパターンは、現在、AppSource のホーム ページをポイントしています。この URL は、AppSource の自分のアドインのページの URL に更新してください。


 <table>
 <tr><th>フィードバックおよび評価</th></tr>
<tr><td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/feedback/office-store"><img src="../images/feedback-rating.png" alt="Feedback and Ratings" style="width: 350px;"/></A></td></tr>
</tr>
 </table>

#### <a name="settings-and-privacy"></a>設定およびプライバシー

アドインには、アドインの動作をユーザーが制御する設定を構成できる設定ページが必要な場合もあります。また、ご自分のアドインが準拠しているプライバシー ポリシーをユーザーに提供したい場合もあるでしょう。 

* **設定**: アドインの動作を制御する構成コンポーネントがある作業ウィンドウを表示します。設定ページには、ユーザーが選択可能なオプションがあります。
    * [設定の仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/patterns/settings.md)
    * [設定のコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)
* **プライバシー ポリシー**: プライバシー ポリシーに関する重要な情報が含まれる作業ウィンドウを表示します。 
    * [プライバシー ポリシーの仕様](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns/blob/master/assets/archived-patterns/general_multiSection.pdf) (この UX 設計パターンはアーカイブされています。 値を評価するときは、この PDF を参照してください)。
    * [プライバシー ポリシーのコード](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings)

<table>
 <tr><th>設定値</th><th>プライバシー ポリシー</th></tr>
<tr>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/settings.png" alt="settings" style="width: 300px;"/></A></td>
<td><A href="https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code/tree/master/templates/settings"><img src="../images/privacy-policy.png" alt="privacy" style="width: 264px;"/></A></td>
</tr></tr>
 </table>

## <a name="see-also"></a>関連項目

* [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
* [Office UI Fabric](http://dev.office.com/fabric/)
