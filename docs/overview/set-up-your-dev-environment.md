---
title: 開発環境をセットアップする
description: 開発者環境をセットアップして、Officeを構築します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: efc89b728117e2888cdebd2c5a132047fe662915
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53774348"
---
# <a name="set-up-your-development-environment"></a>開発環境をセットアップする

このガイドは、クイック スタートまたはチュートリアルに従って、Officeアドインを作成するためのツールをセットアップするのに役立ちます。 以下のリストからツールをインストールする必要があります。 これらが既にインストールされている場合は、クイック スタートなどのクイック スタートを開始するExcel React[準備ができました](../quickstarts/excel-quickstart-react.md)。

- Node.js
- npm
- サブスクリプション Microsoft 365のサブスクリプション バージョンを含むアカウントOffice
- 選択したコード エディター

このガイドでは、コマンド ライン ツールの使い方を知っている必要があります。

## <a name="install-nodejs"></a>Node.js. のインストール

Node.jsは、モダン アドインを開発する必要がある JavaScript ランタイムOfficeです。

Web サイトNode.js最新 [の推奨バージョンをダウンロードしてインストールします](https://nodejs.org)。 オペレーティング システムのインストール手順に従います。

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

ノード バージョン マネージャーを使用して、複数のバージョンの Node.js と npm を切り替える場合がありますが、これは厳密には必要ありません。 これを行う方法の詳細については [、npm の手順を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-microsoft-365"></a>Get Microsoft 365

Microsoft 365 アカウントをまだ持ってない場合は、Microsoft 365 開発者プログラムに参加することで、すべての Office アプリを含む 90 日間の無料のMicrosoft 365 サブスクリプション[を取得できます](https://developer.microsoft.com/office/dev-program)。

## <a name="install-a-code-editor"></a>コード エディターのインストール

以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

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