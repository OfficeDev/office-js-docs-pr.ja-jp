---
title: Office アドインを開発する
description: Office アドイン開発の概要を説明します。
ms.date: 03/11/2022
ms.localizationpriority: high
ms.openlocfilehash: aa56af832d1be3d868700ec4fae731ec55507579
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/22/2022
ms.locfileid: "63711162"
---
# <a name="develop-office-add-ins"></a>Office アドインを開発する

> [!TIP]
> この記事を読む前に、「[Office Add-ins platform overview (Office アドイン プラットフォームの概要)](../overview/office-add-ins.md)」をご覧ください。

すべての Office アドインは、Office アドイン プラットフォーム上で構築します。 どのようなアドインを構築する場合でも、アプリケーションやプラットフォームの可用性、Office JavaScript API のプログラミング パターン、アドインの設定と機能をマニフェスト ファイル上で指定する方法、マニフェストファイルのcapabilities、UIとユーザーエクスペリエンスをデザインする方法など、重要な概念を理解する必要があります。 開発に関するこれらの中心概念については、ドキュメントの **開発ライフサイクル** > **開発** セクションを参照してください。 構築するアドインに対応するアプリケーション固有のドキュメント (たとえば、 [Excel](../excel/index.yml)) を詳しく見る前に、ここに記載される情報を確認してください。

## <a name="create-an-office-add-in"></a>Office アドインを作成する

[Office アドイン用の Yeoman ジェネレーター](yeoman-generator-overview.md) または Visual Studio を使用して Office アドインを作成することができます。

### <a name="yeoman-generator"></a>Yeoman ジェネレーター

Office アドイン用の Yeoman ジェネレーターを使用することで、Visual Studio Code やその他のエディターで管理することができる、Node.js Office アドイン プロジェクトを作成できます。このジェネレーターは、次のいずれに対しても Office アドインを作成できます。

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Excel のカスタム関数

HTML、CSS、JavaScript (または TypeScript) を使用するか、Angular または React を使用してプロジェクトを作成します。 いずれのフレームワークを選択した場合も、JavaScript と Typescript の間から選択することができます。 Yeoman ジェネレーターを使用してアドインを作成する方法については、「[Office アドイン用の Yeoman ジェネレーター](yeoman-generator-overview.md)」を参照してください。

### <a name="visual-studio"></a>Visual Studio

Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。 Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。 Visual Studio を使用してアドインを作成する方法については、「[Visual Studio を使用して Office アドインを開発する](../develop/develop-add-ins-visual-studio.md)」を参照してください。

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understand-the-two-parts-of-an-office-add-in"></a>Office アドインの 2 つの部分について理解する

Office アドインは、2 つの部分から構成されます。

- アドインの設定と機能を定義るアドイン マニフェスト (XML ファイル)。

- 作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドインの UI と機能を定義する Web アプリケーション。

Web アプリケーションでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作します。アドインは、外部 Web サービスの呼び出しやユーザー認証の要求など、Web アプリケーションが一般的に行うその他の機能も実行することができます。

### <a name="define-an-add-ins-settings-and-capabilities"></a>アドインの設定と機能を定義する

Office アドインのマニフェスト (XML ファイル) は、アドインの設定と機能を定義します。 次のような要素を定義するには、マニフェストを構成します。

- アドインを説明するメタデータ (ID、バージョン、説明、表示名、既定のロケールなど)。
- アドインが実行される Office アプリケーション。
- アドインに必要なアクセス許可。
- アドインによって作成されるカスタム UI (カスタム タブ、リボンのボタンなど) などの統合も含めた、アドインの Office との統合方法。
- ブランドおよびコマンドの図像としてアドインで使用される画像の場所。
- アドインの寸法 (例: コンテンツ アドインの寸法、Outlook アドインに対して要求される高さなど)。
- メッセージや予定のコンテキストでアドインをアクティブにさせるタイミングを指定するルール (Outlook アドインのみ)。

マニフェストの詳細については、「[Office アドインの XML マニフェスト](add-in-manifests.md)」を参照してください。

### <a name="interact-with-content-in-an-office-document"></a>Office ドキュメント内のコンテンツを操作する

Office アドインでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメント内のコンテンツを操作できます。

#### <a name="access-the-office-javascript-api-library"></a>Office JavaScript API ライブラリへのアクセス

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>API モデル

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>API 要件セット

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="explore-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Script Lab は、Excel や Word などの Office プログラムでの作業中に Office JavaScript API を調査し、コード スニペットを実行できるようにするアドインです。 これは、[AppSource](https://appsource.microsoft.com/product/office/WA104380862) から無料で利用でき、アドインで必要な機能のプロトタイプを作成したり検証したりする場合に、開発ツールキットに含めておくと便利なツールです。 Script Lab では、組み込みのサンプルのライブラリにアクセスして、簡単に API を試すことができます。また、独自のコードの開始点としてサンプルを使用することもできます。

次の 1 分間のビデオで、Script Lab の実際の動作をご覧ください。

[![Excel、Word、PowerPoint での Script Lab の実行を紹介するショート ビデオ。](../images/screenshot-wide-youtube.png 'Script Lab のプレビュー ビデオ')](https://aka.ms/scriptlabvideo)

Script Lab の詳細については、「[Script Lab を使用して Office JavaScript API を調べる](../overview/explore-with-script-lab.md)」を参照してください。

## <a name="extend-the-office-ui"></a>Office UI をカスタマイズする

Office アドインは、作業ウィンドウ、コンテンツ アドイン、ダイアログ ボックスなど、アドイン コマンドや HTML コンテナーを使用 Office UI を拡張することができます。

- [アドイン コマンド](../design/add-in-commands.md) を使用すると、Office の既定のリボンにカスタム タブ、ボタン、メニューを追加したり、ユーザーが Office ドキュメント内のテキストまたは Excel 内のオブジェクトを右クリックした際に表示される既定のコンテキスト メニューを拡張したりすることができます。 ユーザーがアドイン コマンドを選択すると、アドイン コマンドで指定されているタスク (JavaScript コードの実行、作業ウィンドウを開く、ダイアログ ボックスの起動など) が実行されます。

- [作業ウィンドウ](../design/task-pane-add-ins.md)、[コンテンツ アドイン](../design/content-add-ins.md)、[ダイアログ ボックス](../design/dialog-boxes.md)などの HTML コンテナーを使用すると、カスタム UI を表示させたり Office アプリケーション内で追加機能を表示させたりすることができます。 各作業ウィンドウ、コンテンツ アドイン、またはダイアログ ボックスのコンテンツと機能は、指定した Web ページに由来します。 これらの Web ページでは、Office JavaScript API を使用することで、アドインが実行されている Office ドキュメントのコンテンツを操作できます。また、外部 Web サービスの呼び出しやユーザー認証の要求など、Web ページが一般的に行うその他の機能も実行できます。

次の図では、リボン上に表示されるアドイン コマンド、ドキュメント右側に表示される作業ウィンドウ、およびドキュメント上に表示されるダイアログ ボックスまたはコンテンツ アドインを示しています。

![Office ドキュメントのリボン、タスク ウィンドウ、ダイアログ ボックス / コンテンツ アドイン上のアドイン コマンドを示す図。](../images/add-in-ui-elements.png)

Office UI の拡張とアドインのUXのデザインに関する詳細については、「[Office アドイン用の Office UI 要素](../design/interface-elements.md)」を参照してください。

## <a name="next-steps"></a>次の手順

この記事では、Office アドインの異なる作成方法を説明し、アドインが Office UI を拡張する方法を紹介し、API セットを説明し、Office JavaScript API の探索やアドイン機能のプロトタイプ作成をするための有益なツールとして Script Lab を紹介しました。初歩的な情報の説明は以上になります。Office アドインにの行程を先に進むには、 次の手順を実行してください。

### <a name="create-an-office-add-in"></a>Office アドインを作成する

[5 分間のクイック スタート](../index.yml)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。 以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.yml)を試してみてください。

### <a name="learn-more"></a>詳細情報

Office アドインの開発、テスト、公開の詳細については、このドキュメントを参照してください。

> [!TIP]
> どのようなアドインを構築する場合でも、このドキュメントの 「[開発ライフサイクル](../overview/core-concepts-office-add-ins.md)」セクションに記載する情報に加え、構築するアドインの種類に対応するアプリケーション固有のセクション (たとえば、[Excel](../excel/index.yml)) に記載する情報を使用してください。

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Microsoft 365 開発者プログラムについてご説明します](https://developer.microsoft.com/microsoft-365/dev-program)
- [Office アドインの設計](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)
