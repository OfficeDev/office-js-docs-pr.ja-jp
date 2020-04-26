---
title: ここから開始! 初心者向け Office アドイン開発ガイド
description: Office アドインの学習リソースを使用する初心者向け推奨パス。
ms.date: 04/16/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 026f90ea62960cbbf5ab4420d40a4a9165139cae
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547622"
---
# <a name="start-here-a-guide-for-beginners-making-office-add-ins"></a>ここから開始! 初心者向け Office アドイン開発ガイド

独自のクロスプラットフォーム Office 拡張機能を構築する必要がありますか? 次の手順では、最初に読むべきこと、インストールするツール、完了すべき推奨チュートリアルを示します。

## <a name="step-0-prerequisites"></a>手順 0: 前提条件

- Office アドインは、Office に組み込まれている基本 Web アプリケーションです。 まず、Web アプリケーションの基本について説明し、次に、Web でのホスト方法について説明します。 インターネット、書籍、オンライン コースにこれに関する膨大な情報があります。 Web アプリケーションに関する事前知識がまったくない場合、Bing で "Web アプリとは?" と検索することから始めることを お勧めします。
- Office アドインの作成に使用する主要なプログラミング言語は、JavaScript または TypeScript です。 TypeScript は、JavaScript の厳密に型指定されたバージョンと考えることができます。 これらの言語のいずれにも慣れておらず、VBA、VB.Net、C# の経験がある場合、TypeScript から学習することをお勧めします。 繰り返しになりますが、インターネット、書籍、オンライン コースに、これらの言語に関する豊富な情報があります。

## <a name="step-1-begin-with-fundamentals"></a>手順 1: 基本から始める

今にもコーディングを始めたいと考えておられるかもしれませんが、IDE やコード エディターを開く前に、Office アドインについて、以下をお読みください。

- [Office アドイン プラットフォームの概要](office-add-ins.md): Office Web アドインとは何であるか、VSTO アドインなどの Office を拡張する以前の方法との違いを説明します。
- [Office アドインの構築](office-add-ins-fundamentals.md): ツール、アドイン UI の作成、JavaScript API を使用する Office ドキュメントの操作を含む、Office アドインの開発とライフサイクルの概要を説明します。

これらの記事には多くのリンクが含まれていますが、初心者が Office アドインを使用している場合は、これらを読み、次のセクションに進むときに、ここに戻ることをお勧めします。

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>手順 2: ツールをインストールし、最初のアドインを作成する

全体像を把握できたので、クイック スタートのいずれかを行います。 プラットフォームについて学習する場合は、Excel クイック スタートをお勧めします。 Visual Studio に基づくバージョンがあります。また、node.js と Visual Studio Code に基づくバージョンがあります。

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js および Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>手順 3: コーディング

オーナーズ マニュアルを読んでも、理解することはできません。この [ Excel チュートリアル](../tutorials/excel-tutorial.md)を使用してコーディングを開始してください。 Office JavaScript ライブラリとアドインのマニフェストにある一部の XML を使用します。 後の手順において、両方の背景がわかりやすくなっているため、何も記憶する必要はありません。

## <a name="step-4-understand-the-javascript-library"></a>手順 4: JavaScript ライブラリを理解する

最初に、Microsoft Learn「[Office JavaScript API について理解する](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index)」のこのチュートリアルで、Microsoft Learn ライブラリの全体像を把握します。

次に、API を実行して調査するサンドボックスである [Script Lab ツール](explore-with-script-lab.md)を使用して、Office JavaScript API を学習します。

## <a name="step-5-understand-the-manifest"></a>手順 5: マニフェストを理解する

アドイン マニフェストの目的を理解し、[Office アドイン XML マニフェスト](../develop/add-in-manifests.md)の XML マークアップの概要を理解します。

## <a name="next-steps"></a>次の手順

おめでとうございます。 Office アドインの初心向けラーニング パスを完了しました! ドキュメントの詳細については、以下をご覧ください。

- その他の Office アプリケーション向けのチュートリアルおよびクイック スタート:

  - [OneNote クイック スタート](../quickstarts/onenote-quickstart.md)
  - [Outlook チュートリアル](/outlook/add-ins/addin-tutorial)
  - [PowerPoint チュートリアル](../tutorials/powerpoint-tutorial.md)
  - [Project クイック スタート](../quickstarts/project-quickstart.md)
  - [Word チュートリアル](../tutorials/word-tutorial.md)

- その他の重要な主題:

  - [Office アドインを開発する](../develop/develop-overview.md)
  - [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
  - [Office アドインを設計する](../design/add-in-design.md)
  - [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
  - [Office アドインを展開し、発行する](../publish/publish.md)
  - [リソース](../resources/resources-links-help.md)
