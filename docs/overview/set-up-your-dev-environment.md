---
title: 開発環境をセットアップする
description: Office アドインをビルドするための開発環境をセットアップする
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: af59fb644d1001deb74615d6ced294ad77cbf4e6
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094009"
---
# <a name="set-up-your-development-environment"></a>開発環境をセットアップする

このガイドでは、クイックスタートまたはチュートリアルに従って Office アドインを作成するためのツールのセットアップを支援します。 次の一覧からツールをインストールする必要があります。 これらが既にインストールされている場合は、クイックスタートを開始する準備ができています。たとえば、この[Excel はクイックスタートを反応](../quickstarts/excel-quickstart-react.md)します。

- Node.js
- npm
- Office のサブスクリプション版を含む Microsoft 365 アカウント
- 任意のコードエディター

このガイドでは、コマンドラインツールの使用方法について理解していることを前提としています。 

## <a name="install-nodejs"></a>Node.js. のインストール

Node.js は JavaScript ランタイムで、モダン Office アドインを開発する必要があります。

[Web サイトから最新の推奨バージョンをダウンロード](https://nodejs.org)して、Node.js をインストールします。 オペレーティングシステムのインストール手順に従います。

## <a name="install-npm"></a>Npm をインストールする

npm は、Office アドインの開発に使用されたパッケージをダウンロードするためのオープンソースソフトウェアレジストリです。

Npm をインストールするには、コマンドラインで次のコマンドを実行します。

```command&nbsp;line
    npm install npm -g
```

既に npm がインストールされているかどうかを確認し、インストールされているバージョンを確認するには、コマンドラインで次のコマンドを実行します。

```command&nbsp;line
npm -v
```

ノードバージョンマネージャーを使用して、複数のバージョンの Node.js と npm を切り替えることができますが、これは厳密には必要ありません。 この方法の詳細については、 [「npm の手順」を参照してください](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)。

## <a name="get-office-365"></a>Office 365 を取得する

Microsoft 365 アカウントをまだお持ちでない場合は、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することによって、更新可能な90日間の microsoft 365 サブスクリプションを無料で入手できます。

## <a name="install-a-code-editor"></a>コード エディターのインストール

以下のような Web パーツを構築するのにクライアント側の開発をサポートしている任意のコード エディター、または IDE を使用することができます。

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>次の手順

独自のアドインを作成するか、スクリプトラボを使用して組み込みサンプルを試してみてください。

### <a name="create-an-office-add-in"></a>Office アドインを作成する

[5 分間のクイック スタート](/office/dev/add-ins/)を完了することで、Excel、OneNote、Outlook、PowerPoint、Project、または Word 用の基本的なアドインを簡単に作成することができます。 以前にクイック スタートを完了している場合で、より複雑なアドインを作成したい場合は、[チュートリアル](/office/dev/add-ins/)を試してみてください。

### <a name="explore-the-apis-with-script-lab"></a>Script Lab を使用して API を調べる

Office JavaScript API でどのような機能が提供されているかを把握するには、[Script Lab](explore-with-script-lab.md) に組み込まれているサンプルのライブラリを参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインを構築する](../overview/office-add-ins-fundamentals.md)
- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)
