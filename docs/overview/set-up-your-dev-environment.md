---
title: 開発環境をセットアップする
description: 新しいアドインを構築するためのOfficeをセットアップします。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 1dd0cc6bb035a0274e36fe9916dcd2481bdf0b39
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234129"
---
# <a name="set-up-your-development-environment"></a>開発環境をセットアップする

このガイドは、クイック スタートまたはチュートリアルに従Officeアドインを作成するためのツールをセットアップする場合に役立ちます。 次の一覧からツールをインストールする必要があります。 既にインストールされている場合は、この Excel React クイック スタートなど、クイック スタートを [開始する準備が整っています](../quickstarts/excel-quickstart-react.md)。

- Node.js
- npm
- サブスクリプション バージョンを含む Microsoft 365 Office
- 選択したコード エディター

このガイドでは、コマンド ライン ツールの使い方を知っている必要があります。 

## <a name="install-nodejs"></a>Node.js. のインストール

Node.jsは、モダン アドインを開発するために必要Office JavaScript ランタイムです。

Web Node.js [から最新の推奨バージョンをダウンロードしてインストールします](https://nodejs.org)。 オペレーティング システムのインストール手順に従います。

## <a name="install-npm"></a>npm をインストールする

npm はオープン ソース のソフトウェア レジストリで、アドインの開発に使用されるパッケージOfficeダウンロードできます。

npm をインストールするには、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
    npm install npm -g
```

npm が既にインストール済みで、インストールされているバージョンを確認するには、コマンド ラインで次のコマンドを実行します。

```command&nbsp;line
npm -v
```

ノード バージョン マネージャーを使用して、複数のバージョンの Node.js と npm を切り替える場合がありますが、これは厳密には必要ありません。 これを行う方法の詳細については [、npm の手順を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-microsoft-365"></a>Microsoft 365 を取得する

まだ Microsoft 365 アカウントをお持ちでない場合は、Microsoft 365 開発者プログラムに参加することで、すべての Office アプリを含む 90 日間の更新可能な [無料の Microsoft 365](https://developer.microsoft.com/office/dev-program)サブスクリプションを取得できます。

## <a name="install-a-code-editor"></a>コード エディターのインストール

以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>次の手順

独自のアドインを作成するか、Script Lab を使用して組み込みのサンプルを試してみてください。

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