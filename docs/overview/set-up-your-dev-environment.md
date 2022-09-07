---
title: 開発環境をセットアップする
description: Office アドインをビルドするように開発環境を設定します。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e03ea7f55786107354f9d5a92e0cb30ffb559ec
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616006"
---
# <a name="set-up-your-development-environment"></a>開発環境をセットアップする

このガイドは、クイック スタートまたはチュートリアルに従って Office アドインを作成できるようにツールを設定するのに役立ちます。 既にインストールされている場合は、この [Excel Reactクイック スタートなどのクイック スタート](../quickstarts/excel-quickstart-react.md)を開始する準備が整います。

## <a name="get-microsoft-365"></a>Microsoft 365 を入手する

Microsoft 365 アカウントが必要です。 Microsoft 365 開発者プログラムに参加することで、すべての Office アプリを含む 90 日間の無料の更新可能な [Microsoft 365](https://developer.microsoft.com/office/dev-program) サブスクリプションを入手できます。

## <a name="install-the-environment"></a>環境をインストールする

選択できる開発環境は 2 種類あります。 2 つの環境で作成される Office アドイン プロジェクトのスキャフォールディングは異なるため、複数のユーザーがアドイン プロジェクトで作業する場合は、すべて同じ環境を使用する必要があります。 

- **Node.js環境**: 推奨されます。 この環境では、ツールがインストールされ、コマンド ラインで実行されます。 アドインの Web アプリケーションパーツのサーバー側は JavaScript または TypeScript で記述され、Node.js ランタイムでホストされます。 この環境には、Office リンターや WebPack というバンドルャー/タスクランナーなど、多くの便利なアドイン開発ツールがあります。 プロジェクトの作成とスキャフォールディング ツールである Yo Office は頻繁に更新されます。
- **Visual Studio 環境**: 開発用コンピューターが Windows であり、.NET ベースの言語とフレームワーク (ASP.NET など) を使用してアドインのサーバー側を開発する場合にのみ、この環境を選択します。 Visual Studio のアドイン プロジェクト テンプレートは、Node.js環境のテンプレートほど頻繁に更新されません。 組み込みの Visual Studio デバッガーではクライアント側コードをデバッグできませんが、ブラウザーの開発ツールを使用してクライアント側コードをデバッグできます。 **Visual Studio 環境** タブで後ほど詳しく説明します。

> [!NOTE]
> Visual Studio for Mac Office アドイン用のプロジェクト スキャフォールディング テンプレートは含まれていないため、開発用コンピューターが Mac の場合は、Node.js環境で作業する必要があります。

選択した環境のタブを選択します。 

# <a name="nodejs-environment"></a>[Node.js環境](#tab/yeomangenerator)

インストールする主なツールは次のとおりです。

- Node.js
- npm
- 任意のコード エディター
- Yo Office
- Office JavaScript のリンター

このガイドでは、コマンド ライン ツールの使用方法を理解していることを前提としています。

### <a name="install-nodejs-and-npm"></a>Node.jsと npm をインストールする

Node.jsは、最新の Office アドインの開発に使用する JavaScript ランタイムです。

[web サイトから最新の推奨バージョンをダウンロードして、Node.jsをインストールします](https://nodejs.org)。 オペレーティング システムのインストール手順に従います。

npm は、Office アドインの開発に使用されるパッケージをダウンロードするオープンソースソフトウェア レジストリです。通常、Node.jsをインストールすると自動的にインストールされます。 npm が既にインストールされていて、インストールされているバージョンが表示されているかどうかを確認するには、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
npm -v
```

何らかの理由で手動でインストールする場合は、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> Node バージョン マネージャーを使用して、複数のバージョンのNode.jsと npm を切り替えることができますが、これは厳密には必要ありません。 これを行う方法の詳細については、 [npm の手順を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

### <a name="install-a-code-editor"></a>コード エディターのインストール

以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。

- [Visual Studio Code](https://code.visualstudio.com/) (推奨)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>Yeoman ジェネレーター &mdash; Yo Office をインストールする

プロジェクトの作成とスキャフォールディング ツールは、 [Office アドイン用の Yeoman ジェネレーターです](../develop/yeoman-generator-overview.md)。これは、 **Yo Office** と呼ばれます。 [Yeoman](https://github.com/yeoman/yo) と Yo Office の最新バージョンをインストールする必要があります。 以上のツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>Office JavaScript リンターをインストールして使用する

Microsoft では、Office JavaScript ライブラリを使用するときに一般的なエラーをキャッチするのに役立つ JavaScript リンターを提供しています。 リンターをインストールするには、次の 2 つのコマンドを実行します ( [Node.jsと npm をインストール](#install-nodejs-and-npm)した後)。

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Office アドインツール [用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md) を使用して Office アドイン プロジェクトを作成する場合は、残りのセットアップが自動的に行われます。 Visual Studio Code などのエディターのターミナルで、またはコマンド プロンプトで、次のコマンドを使用してリンターを実行します。 リンターによって見つかった問題は、ターミナルまたはプロンプトに表示され、Visual Studio Code などのリンター メッセージをサポートするエディターを使用している場合にも、コードに直接表示されます。 (Yeoman ジェネレーターのインストールの詳細については、「 [Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)」を参照してください)。

```command&nbsp;line
npm run lint
```

アドイン プロジェクトが別の方法で作成された場合は、次の手順に従います。

1. プロジェクトのルートで、 **.eslintrc.json** という名前のテキスト ファイルを作成します (まだ存在しない場合)。 名前付きの `plugins` プロパティと `extends`、両方の型配列があることを確認します。 配列に `plugins` 含める `"office-addins"` 必要があり、配列に `extends` 含める `"plugin:office-addins/recommended"`必要があります。 次に簡単な例を示します。 **.eslintrc.json** ファイルには、追加のプロパティと 2 つの配列の追加メンバーが含まれている場合があります。

   ```json
   {
     "plugins": [
       "office-addins"
     ],
     "extends": [
       "plugin:office-addins/recommended"
     ]
   }
   ```

1. プロジェクトのルートで **package.json** ファイルを開き、配列に次の `scripts` メンバーがあることを確認します。

   ```json
   "lint": "office-addin-lint check",
   ```

1. Visual Studio Code などのエディターのターミナルで、またはコマンド プロンプトで、次のコマンドを使用してリンターを実行します。 リンターによって見つかった問題は、ターミナルまたはプロンプトに表示され、Visual Studio Code などのリンター メッセージをサポートするエディターを使用している場合にも、コードに直接表示されます。

   ```command&nbsp;line
   npm run lint
   ```

# <a name="visual-studio-environment"></a>[Visual Studio 環境](#tab/visualstudio)

### <a name="install-visual-studio"></a>Visual Studio のインストール

Visual Studio 2017 (Windows 用) 以降がインストールされていない場合は、 [Visual Studio ダウンロード](https://visualstudio.microsoft.com/downloads/)から最新バージョンをインストールします。 インストーラーからワークロードの指定を求められた場合は、 **Office/SharePoint 開発** ワークロードを必ず含めます。 必要になる可能性があるその他のワークロードは、.NET、**JavaScript、TypeScript 言語のサポート** (アドインのクライアント側をコーディングするための)、および ASP.NET 関連のワークロード **用の Web 開発ツール** です。

> [!TIP]
> 2022 年 6 月現在、Visual Studio と共にインストールされている Office アドイン マニフェストの XML スキーマは最新バージョンではありません。 これは、アドインが使用するアドイン機能によっては、アドインに影響を与える可能性があります。 そのため、マニフェストの XML スキーマを更新する必要がある場合があります。 詳細については、「 [Visual Studio プロジェクトのマニフェスト スキーマ検証エラー](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)」を参照してください。

> [!NOTE]
> Visual Studio 環境を使用している場合のクライアント側コードのデバッグについては、「 [Visual Studio での Office アドインのデバッグ](../develop/debug-office-add-ins-in-visual-studio.md)」を参照してください。 Visual Studio で作成した Web アプリケーションと同じように、サーバー側のコードをデバッグします。 [クライアント側またはサーバー側を](../testing/debug-add-ins-overview.md#server-side-or-client-side)参照してください。

---

## <a name="install-script-lab"></a>Script Labをインストールする

Script Labは、Office JavaScript ライブラリ API を呼び出すコードをすばやくプロトタイプ作成するためのツールです。 Script Lab自体は Office アドインであり、AppSource から[インストール](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1)Script Lab。 Excel、PowerPoint、Word のバージョンと、Outlook 用の別のバージョンがあります。 Script Labを使用する方法については、「Script Labを[使用して Office JavaScript API を探索](explore-with-script-lab.md)する」を参照してください。

## <a name="next-steps"></a>次の手順

独自のアドインを作成するか[、Script Lab](explore-with-script-lab.md)を使用して組み込みのサンプルを試してみてください。

### <a name="create-an-office-add-in"></a>Office アドインを作成する

[5 分間のクイック スタート](../index.yml)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。 以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.yml)を試してみてください。

### <a name="explore-the-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインの設計](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)
- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)