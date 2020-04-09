---
title: Office アドインの構築
description: Office アドイン開発の概要を説明します。
ms.date: 02/27/2020
localization_priority: Priority
ms.openlocfilehash: 4b9e6b6e9f282ca4dc4831f2dcdda88f379e6c33
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185611"
---
# <a name="building-office-add-ins"></a>Office アドインの構築

> [!TIP]
> この記事を読む前に、「[Office Add-ins platform overview (Office アドイン プラットフォームの概要)](office-add-ins.md)」をご覧ください。

Office アドインは、Office アプリケーションの UI と機能を拡張し、Office ドキュメント内のコンテンツを操作します。 Word、Excel、PowerPoint、OneNote、Project、Outlook の拡張と操作を行うアドインの構築には、一般的な Web テクノロジを使用します。 構築するアドインは、Windows、Mac、iPad やブラウザー上など、複数のプラットフォーム上の Office で実行できます。 この記事では、Office アドイン開発の概要を説明します。

## <a name="creating-an-office-add-in"></a>Office アドインの作成 

Office アドイン用の Yeoman ジェネレーターまたは Visual Studio を使用して Office アドインを作成することができます。

### <a name="yeoman-generator-for-office-add-ins"></a>Office アドイン用の Yeoman ジェネレーター

[Office アドイン用の Yeoman ジェネレーター](https://github.com/officedev/generator-office)を使用することで、Visual Studio Code やその他のエディターで管理することができる、Node.js Office アドイン プロジェクトを作成できます。 ジェネレーターでは、次のいずれのホスト用の Office アドインも作成できます。

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Excel のカスタム関数

プロジェクトを作成するのに、HTML、CSS、および JavaScript を使用するのか、Angular または React を使用するのかを選択できます。 いずれのフレームワークを選択した場合も、JavaScript と Typescript の間から選択することができます。 Yeoman ジェネレーターを使用してアドインを作成する方法については、「[Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)」を参照してください。

### <a name="visual-studio"></a>Visual Studio

Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。 Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。 Visual Studio を使用してアドインを作成する方法については、「[Visual Studio を使用して Office アドインを開発する](../develop/develop-add-ins-visual-studio.md)」を参照してください。

[!include[Yeoman vs Visual Studio comparision](../includes/yeoman-generator-recommendation.md)]

## <a name="exploring-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Script Lab は、Excel や Word などの Office プログラムでの作業中に Office JavaScript API を調査し、コード スニペットを実行できるようにするアドインです。 これは、[AppSource](https://appsource.microsoft.com/product/office/WA104380862) から無料で利用でき、アドインで必要な機能のプロトタイプを作成したり検証したりする場合に、開発ツールキットに含めておくと便利なツールです。 Script Lab では、組み込みのサンプルのライブラリにアクセスして、簡単に API を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。 

次の 1 分間のビデオで、Script Lab の実際の動作をご覧ください。

[![Excel、Word、PowerPoint での Script Lab の実行を紹介するプレビュー ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)

Script Lab の詳細については、「[Script Lab を使用して Office JavaScript API を調べる](../overview/explore-with-script-lab.md)」を参照してください。

## <a name="extending-the-office-ui"></a>Office UI の拡張

Office アドインは、作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドイン コマンドや HTML コンテナーを使用 Office UI を拡張することができます。

- [アドイン コマンド](../design/add-in-commands.md) を使用すると、Office の既定のリボンにカスタム タブ、ボタン、メニューを追加したり、ユーザーが Office ドキュメント内のテキストまたは Excel 内のオブジェクトを右クリックした際に表示される既定のコンテキスト メニューを拡張したりすることができます。 ユーザーがアドイン コマンドを選択すると、アドイン コマンドで指定されているタスク (JavaScript コードの実行、作業ウィンドウを開く、ダイアログ ボックスの起動など) が実行されます。

- [作業ウィンドウ](../design/task-pane-add-ins.md)、[コンテンツ アドイン](../design/content-add-ins.md)、[ダイアログ ボックス](../design/dialog-boxes.md)などの HTML コンテナーを使用すると、カスタム UI を表示させたり Office アプリケーション内で追加機能を表示させたりすることができます。 各作業ウィンドウ、コンテンツ アドイン、またはダイアログ ボックスのコンテンツと機能は、指定した Web ページに由来します。 これらの Web ページでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメントのコンテンツを操作できます。また、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行できます。

次の図では、リボン上に表示されるアドイン コマンド、ドキュメント右側に表示される作業ウィンドウ、およびドキュメント上に表示されるダイアログ ボックスまたはコンテンツ アドインを示しています。

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス上のアドイン コマンドを示す図](../images/add-in-ui-elements.png)

Office UI の拡張に関する詳細については、「[Office アドイン用の Office UI 要素](../design/interface-elements.md)」を参照してください。

## <a name="core-development-concepts"></a>開発の中心概念 

Office アドインは、2 つの部分から構成されます。

- アドインの設定と機能を定義るアドイン マニフェスト (XML ファイル)。

- 作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドインの UI と機能を定義する Web アプリケーション。

Web アプリケーションでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作します。 アドインは、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行することができます。

### <a name="defining-an-add-ins-settings-and-capabilities"></a>アドインの設定と機能を定義する

Office アドインのマニフェスト (XML ファイル) は、アドインの設定と機能を定義します。 次のような要素を定義するには、マニフェストを構成します。

- アドインを説明するメタデータ (ID、バージョン、説明、表示名、既定のロケールなど)。
- アドインが実行される Office アプリケーション。
- アドインに必要なアクセス許可。
- アドインによって作成されるカスタム UI (カスタム タブ、リボンのボタンなど) などの統合も含めた、アドインの Office との統合方法。
- ブランドおよびコマンドの図像としてアドインで使用される画像の場所。
- アドインの寸法 (例: コンテンツ アドインの寸法、Outlook アドインに対して要求される高さなど)。
- メッセージや予定のコンテキストでアドインをアクティブにさせるタイミングを指定するルール (Outlook アドインのみ)。

マニフェストの詳細については、「[Office アドインの XML マニフェスト](add-in-manifests.md)」を参照してください。

### <a name="interacting-with-content-in-an-office-document"></a>Office ドキュメント内のコンテンツを操作する

Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。 

#### <a name="accessing-the-office-javascript-api-library"></a>Office JavaScript API ライブラリへのアクセス

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>API モデル

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>API 要件セット

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

## <a name="testing-and-debugging-an-office-add-in"></a>Office アドインのテストとデバッグ

アドインの開発中は、_サイドロード_という手法を使用してアドインをローカルでテストできます。 アドインをサイドロードする手順はプラットフォームによって異なり、一部のケースでは、製品ごとに異なります。 同様に、アドインのデバッグ手順も、プラットフォームや製品によって異なります。 テストとデバッグの詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。

## <a name="publishing-an-office-add-in"></a>Office アドインの公開

アドインを他のユーザーと共有する準備ができたら、目的に一番合った展開方法を使用してアドインを共有します。 たとえば、組織内のユーザーにアドインを展開する場合は、一元展開を使用するか、アドインを SharePoint アプリ カタログで公開することをお勧めします。 すべてのユーザーが入手できるようにアドインを一般公開する場合は、アドインを AppSource で公開できます。 公開の詳細については、「[Office アドインの展開と公開](../publish/publish.md)」を参照してください。

## <a name="next-steps"></a>次のステップ

この記事では、Office アドインの異なる作成方法を説明し、Office JavaScript API の調査とアドイン機能のプロトタイプ作成における効果的なツールとして Script Lab を紹介し、Office アドインの開発、テスト、および公開に関する重要な概念の説明を行いました。 初歩的な情報の説明は以上になります。Office アドインにの行程を先に進むには、 次の手順を実行してください。

### <a name="create-an-office-add-in"></a>Office アドインを作成する

[5 分間のクイック スタート](../index.md)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。 以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.md)を試してみてください。

### <a name="explore-the-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。

### <a name="learn-more"></a>詳細情報

Office アドインの開発、テスト、公開の詳細については、このドキュメントを参照してください。

> [!TIP]
> どのようなアドインを構築する場合でも、このドキュメントの 「[中心概念](core-concepts-office-add-ins.md)」セクションに記載する情報に加え、構築するアドインの種類に対応するホスト固有のセクション (たとえば、[Excel](../excel/index.md)) に記載する情報を使用してください。
>
> ![目次を表示する画像](../images/top-level-toc.png)

## <a name="see-also"></a>関連項目 

- [Office アドイン プラットフォームの概要](office-add-ins.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Visual Studio Code を使用して Office アドインを開発する](../develop/develop-add-ins-vscode.md)
- [Visual Studio を使用して Office アドインを開発する](../develop/develop-add-ins-visual-studio.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)