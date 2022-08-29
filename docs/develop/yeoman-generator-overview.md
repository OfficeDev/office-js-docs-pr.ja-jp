---
title: Yeoman ジェネレーターを使用して Office アドイン プロジェクトを作成する
description: Office アドイン用 Yeoman ジェネレーターを使用して Office アドイン プロジェクトを作成する方法について説明します。
ms.date: 08/19/2022
ms.localizationpriority: high
ms.openlocfilehash: f109c4dbc386a4cc23f72d0c67f9e4904360bba4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422797"
---
# <a name="create-office-add-in-projects-using-the-yeoman-generator"></a>Yeoman ジェネレーターを使用して Office アドイン プロジェクトを作成する

[Office アドイン用 Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office) ("Yo Office" とも呼ばれます) は、Office アドイン開発プロジェクトを作成する対話型のNode.js ベースのコマンド ライン ツールです。 このツールを使用してアドイン プロジェクトを作成することをお勧めします。ただし、アドインのサーバー側コードを .NET ベースの言語 (C# や VB.Net など) またはアドインをインターネット インフォメーション サーバー (IIS) でホストする必要があります。 後者の 2 つの状況のいずれかで、 [Visual Studio を使用してアドインを作成します](develop-add-ins-visual-studio.md)。

ツールが作成するプロジェクトには、次の特性があります。

- **これらは、package.json** ファイルを含む標準 [の npm](https://www.npmjs.com/) 構成を持っています。
- プロジェクトのビルド、サーバーの起動、アドインの Office でのサイドロード、その他のタスクに役立つスクリプトがいくつか含まれています。
- バンドルャーと基本的なタスク ランナーとして [Webpack](https://webpack.js.org/) を使用します。
- 開発モードでは、ホットリロードと再コンパイルオン変更をサポートする開発指向の [高速](http://expressjs.com/) サーバーである webpack のNode.jsベースの webpack-dev-server によって localhost でホストされます。
- 既定では、すべての依存関係はツールによってインストールされますが、コマンド ライン引数を使用してインストールを延期できます。
- 完全なアドイン マニフェストが含まれています。
- "Hello World" レベルのアドインがあり、ツールが完了するとすぐに実行できます。
- これには、TypeScript と最新バージョンの JavaScript を ES5 JavaScript にトランスパイルするように構成されたポリフィルとトランスパイラーが含まれます。 これらの機能により、Internet Explorer を含め、Office アドインが実行される可能性があるすべてのランタイムでアドインが確実にサポートされます。

> [!TIP]
> 別のタスク ランナーや別のサーバーを使用するなど、これらの選択肢から大きく逸脱したい場合は、ツールを実行するときに [マニフェスト専用オプション](#manifest-only-option)を選択することをお勧めします。

## <a name="install-the-generator"></a>ジェネレーターをインストールする

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="use-the-tool"></a>ツールを使用する

システム プロンプトで次のコマンドを使用してツールを起動します (bash ウィンドウではありません)。

```command&nbsp;line
yo office 
```

多くの読み込みが必要なため、ツールの起動までに 20 秒かかる場合があります。 このツールでは、一連の質問が表示されます。 一部のユーザーは、プロンプトに対する回答を入力するだけです。 他のユーザーには、考えられる回答の一覧が表示されます。 指定されたリストがある場合は、1 つを選択し、Enter キーを押します。

最初の質問では、6 種類のプロジェクトから選択するように求められます。 以下のオプションがあります:

- **Office アドイン作業ウィンドウ プロジェクト**
- **Angular フレームワークを使用した Office アドイン作業ウィンドウ プロジェクト**
- **React フレームワークを使用した Office アドイン作業ウィンドウ プロジェクト**
- **シングル サインオンをサポートする Office アドイン作業ウィンドウ プロジェクト**
- **マニフェストのみを含む Office アドイン プロジェクト**
- **Excel カスタム関数アドイン プロジェクト**

![Yeoman ジェネレーターのプロジェクトの種類と可能な回答のプロンプトを示すスクリーンショット。](../images/yo-office-project-type-prompt.png)

> [!NOTE]
> **シングル サインオン オプションをサポートする Office アドイン作業ウィンドウ プロジェクト** では、アドインでのシングル サインオン (SSO) のしくみを確認するために使用できるプロジェクトが生成されます。 プロジェクトを運用アドインの基礎として使用することはできません。 運用アドインの基礎となる SSO 対応プロジェクトを取得するには、 [サンプル リポジトリの SSO サンプルの 1 つの](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth) "完了" バージョンを参照してください。
>
> **マニフェストのみのオプションを含む Office アドイン プロジェクト** では、基本的なアドイン マニフェストと最小限のスキャフォールディングを含むプロジェクトが生成されます。 このオプションの詳細については、「 [マニフェスト専用オプション](#manifest-only-option)」を参照してください。

次の質問では、 **TypeScript と JavaScript** の間で選択するように求 **められます**。 (前の質問でマニフェストのみのオプションを選択した場合、この質問はスキップされます)。

![ユーザーが上記の質問に対して [Office アドイン作業ウィンドウ プロジェクト] を選択したことを示すスクリーンショット。Yeoman ジェネレーターで言語のプロンプトと、考えられる回答 (TypeScript と JavaScript) が表示されます。](../images/yo-office-language-prompt.png)

次に、アドインに名前を付けるメッセージが表示されます。 指定した名前はアドインのマニフェストで使用されますが、後で変更することもできます。

![ユーザーが前の質問に TypeScript を選択し、Yeoman ジェネレーターでアドイン名のプロンプトを示すスクリーンショット。](../images/yo-office-name-prompt.png)

次に、アドインを実行する Office アプリケーションを選択するように求められます。 **Excel**、**OneNote**、**Outlook**、**PowerPoint**、**Project**、**Word** の 6 つのアプリケーションから選択できます。 1 つだけ選択する必要がありますが、後でマニフェストを変更して、追加の Office アプリケーションをサポートすることができます。 例外は Outlook です。 Outlook をサポートするマニフェストは、他の Office アプリケーションをサポートできません。

![ユーザーがプロジェクトに "Contoso アドイン" という名前を付け、Yeoman ジェネレーターで Office アプリケーションのプロンプトと可能な回答を示すスクリーンショット。](../images/yo-office-host-prompt.png)

この質問に答えたら、ジェネレーターによってプロジェクトが作成され、依存関係がインストールされます。 画面の npm 出力に **WARN** メッセージが表示される場合があります。 これらは無視できます。 また、脆弱性が見つかったというメッセージが表示される場合もあります。 現時点ではこれらを無視できますが、アドインを運用環境にリリースする前に、最終的に修正する必要があります。 脆弱性の修正の詳細については、ブラウザーを開き、"npm の脆弱性" を検索してください。

作成が成功した場合は、 **おめでとうございます。** メッセージをコマンド ウィンドウに表示し、次に推奨される手順をいくつか示します。 (クイック スタートまたはチュートリアルの一部としてジェネレーターを使用している場合は、コマンド ウィンドウの次の手順を無視して、記事の手順を続行します)。

> [!TIP]
> Office アドイン プロジェクトのスキャフォールディングを作成するが、依存関係のインストールを延期する場合は、コマンドに `--skip-install` オプションを `yo office` 追加します。 以下にコードの例を示します。
>
> ```command&nbsp;line
> yo office --skip-install
> ```
>
> 依存関係をインストールする準備ができたら、コマンド プロンプトでプロジェクトのルート フォルダーに移動し、次のように入力 `npm install`します。

## <a name="manifest-only-option"></a>マニフェスト専用オプション

このオプションを使用すると、アドインのマニフェストのみが作成されます。 結果のプロジェクトには、Hello World アドイン、スクリプト、依存関係はありません。 次のシナリオでは、このオプションを使用します。

- Yeoman ジェネレーター プロジェクトが既定でインストールおよび構成するツールとは異なるツールを使用する必要があります。 たとえば、別のバンドルャー、トランスパイラー、タスク ランナー、または開発サーバーを使用します。
- Vue などのAngularやReact以外の Web アプリケーション開発フレームワークを使用する必要があります。

マニフェストのみのオプションでジェネレーターを使用する例については、「 [Vue を使用して Excel 作業ウィンドウ アドインをビルドする](../quickstarts/excel-quickstart-vue.md)」を参照してください。

## <a name="use-command-line-parameters"></a>コマンド ライン パラメーターを使用する

コマンドにパラメーターを `yo office` 追加することもできます。 2 つの最も一般的なオプションは、次のとおりです。

- `yo office --details`: これにより、他のすべてのコマンド ライン パラメーターに関する簡単なヘルプが出力されます。
- `yo office --skip-install`: これにより、ジェネレーターが依存関係をインストールできなくなります。

コマンド ライン パラメーターの詳細については、 [Office アドイン用 Yeoman ジェネレーターのジェネレーター](https://github.com/officedev/generator-office)の readme を参照してください。

## <a name="troubleshooting"></a>トラブルシューティング

ツールの使用に問題が発生した場合は、最初に再インストールして、最新バージョンを使用していることを確認する必要があります。 (詳細については、「 [ジェネレーターのインストール](#install-the-generator) 」を参照してください)。問題が解決しない場合は、 [GitHub リポジトリの問題を](https://github.com/OfficeDev/generator-office/issues) 検索して、他のユーザーが同じ問題に遭遇し、解決策が見つかったかどうかを確認します。 誰もいない場合は、 [新しい問題を作成します](https://github.com/OfficeDev/generator-office/issues/new?assignees=&labels=needs+triage&template=bug_report.md&title=)。
