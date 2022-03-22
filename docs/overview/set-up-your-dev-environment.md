---
title: 開発環境をセットアップする
description: 開発者環境をセットアップして、Officeを構築します。
ms.date: 10/26/2021
ms.localizationpriority: medium
ms.openlocfilehash: ad1fc265640b6fb5931ba2086cc61784e94365c1
ms.sourcegitcommit: 4a7b9b9b359d51688752851bf3b41b36f95eea00
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/22/2022
ms.locfileid: "63710924"
---
# <a name="set-up-your-development-environment"></a>開発環境をセットアップする

このガイドは、クイック スタートまたはチュートリアルに従って、Officeアドインを作成するためのツールをセットアップするのに役立ちます。 以下のリストからツールをインストールする必要があります。 これらが既にインストールされている場合は、クイック スタートなどのクイック スタートを開始Excel React[準備が整いました](../quickstarts/excel-quickstart-react.md)。

- Node.js
- npm
- サブスクリプション Microsoft 365のサブスクリプション バージョンを含むアカウントOffice
- 選択したコード エディター
- JavaScript Officeインター

このガイドでは、コマンド ライン ツールの使い方を知っている必要があります。

## <a name="install-nodejs"></a>Node.js. のインストール

Node.jsは、モダン アドインを開発する必要がある JavaScript ランタイムOfficeです。

Web サイトNode.js [最新の推奨バージョンをダウンロードしてインストールします](https://nodejs.org)。 オペレーティング システムのインストール手順に従います。

## <a name="install-npm"></a>npm のインストール

npm は、アドインの開発に使用されるパッケージをダウンロードするOfficeソフトウェア レジストリです。

npm をインストールするには、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
    npm install npm -g
```

npm が既にインストールされていることを確認し、インストールされているバージョンを確認するには、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
npm -v
```

ノード バージョン マネージャーを使用して、複数のバージョンの Node.js と npm を切り替える場合がありますが、これは厳密には必要ありません。 これを行う方法の詳細については、 [npm の手順を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-microsoft-365"></a>Get Microsoft 365

Microsoft 365 アカウントをまだ持ってない場合は、Microsoft 365 開発者プログラムに参加することで、すべての Office アプリを含む 90 日間の無料のMicrosoft 365 サブスクリプション[を](https://developer.microsoft.com/office/dev-program)取得できます。

## <a name="install-a-code-editor"></a>コード エディターのインストール

以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="install-and-use-the-office-javascript-linter"></a>JavaScript linter のOffice使用する

Microsoft では、JavaScript ライブラリを使用するときに一般的なエラーをキャッチするのに役立つ JavaScript Office提供されています。 linter をインストールするには、次の 2 つのコマンドを実行します (Node.js[ ](#install-nodejs) [npm をインストールした](#install-npm)後)。

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Office アドイン ツール用の [Yeoman](../develop/yeoman-generator-overview.md) ジェネレーターを使用して Office アドイン プロジェクトを作成すると、セットアップの残りの部分が実行されます。 次のコマンドを使用して、エディターのターミナル (コマンド プロンプトなど) で linter をVisual Studio Codeコマンド プロンプトで実行します。 linter で見つかった問題は、ターミナルまたはプロンプトに表示され、Visual Studio Code などの linter メッセージをサポートするエディターを使用している場合にも、コードに直接表示されます。 (Yeoman ジェネレーターのインストールの詳細については、「[Yeoman generator for Officeアドイン」を参照](../develop/yeoman-generator-overview.md)してください。

```command&nbsp;line
npm run lint
```

アドイン プロジェクトが別の方法で作成された場合は、次の手順を実行します。

1. プロジェクトのルートで、 **.eslintrc.json** という名前のテキスト ファイルを作成します (まだ存在しない場合)。 配列型のプロパティと、 `plugins` 両方の型配列 `extends`を持つ必要があります。 配列`plugins`は含める必要があります。`"office-addins"`配列には`extends`.`"plugin:office-addins/recommended"` 次に簡単な例を示します。 **.eslintrc.json** ファイルには、2 つの配列の追加のプロパティと追加のメンバーが含まれます。

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

1. プロジェクトのルートで **package.json** ファイルを開き、 `scripts` 配列に次のメンバーが含まれています。

   ```json
   "lint": "office-addin-lint check",
   ```

1. 次のコマンドを使用して、エディターのターミナル (コマンド プロンプトなど) で linter をVisual Studio Codeコマンド プロンプトで実行します。 linter で見つかった問題は、ターミナルまたはプロンプトに表示され、Visual Studio Code などの linter メッセージをサポートするエディターを使用している場合にも、コードに直接表示されます。

   ```command&nbsp;line
   npm run lint
   ```

## <a name="next-steps"></a>次の手順

独自のアドインを作成するか、Script Labを使用して組み込みのサンプルを試してみてください。

### <a name="create-an-office-add-in"></a>Office アドインを作成する

[5 分間のクイック スタート](../index.yml)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。 以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](../index.yml)を試してみてください。

### <a name="explore-the-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)
- [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)