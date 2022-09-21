---
title: Office アドイン入門
description: Office アドインの学習リソースを使用する初心者向け推奨パス。
ms.date: 02/12/2021
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 097be9f7aa1563dc513da9cb27eeb7daa344aa41
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857543"
---
# <a name="beginners-guide"></a>初心者ガイド

独自のクロスプラットフォーム Office 拡張機能を構築する必要がありますか? 次の手順では、最初に読むべきこと、インストールするツール、完了すべき推奨チュートリアルを示します。

> [!NOTE]
> Office 用の VSTO アドインの作成経験がある場合には、この記事内にある情報のスーパーセットである「[VSTO アドイン作成者のためのガイド](learning-path-transition.md)」を今すぐご覧になることをお勧めします。

## <a name="step-0-prerequisites"></a>手順 0: 前提条件

- Office アドインは、Office に組み込まれている基本 Web アプリケーションです。 まず、Web アプリケーションの基本について説明し、次に、Web でのホスト方法について説明します。 インターネット、書籍、オンライン コースにこれに関する膨大な情報があります。 Web アプリケーションに関する事前知識がまったくない場合、Bing で "Web アプリとは?" と検索することから始めることを お勧めします。
- Office アドインの作成に使用する主要なプログラミング言語は、JavaScript または TypeScript です。 TypeScript は、JavaScript の厳密に型指定されたバージョンと考えることができます。 これらの言語のいずれにも慣れておらず、VBA、VB.Net、C# の経験がある場合、TypeScript から学習することをお勧めします。 繰り返しになりますが、インターネット、書籍、オンライン コースに、これらの言語に関する豊富な情報があります。

## <a name="step-1-begin-with-fundamentals"></a>手順 1: 基本から始める

今にもコーディングを始めたいと考えておられるかもしれませんが、IDE やコード エディターを開く前に、Office アドインについて、以下をお読みください。

- [Office アドイン プラットフォームの概要](office-add-ins.md): Office Web アドインとは何であるか、VSTO アドインなどの Office を拡張する以前の方法との違いを説明します。
- [Office アドインを開発する](../develop/develop-overview.md): ツール、アドイン UI の作成、JavaScript API を使用する Office ドキュメントの操作を含む、Office アドインの開発とライフサイクルの概要を説明します。

これらの記事には多くのリンクが含まれていますが、初心者が Office アドインを使用している場合は、これらを読み、次のセクションに進むときに、ここに戻ることをお勧めします。

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>手順 2: ツールをインストールし、最初のアドインを作成する

全体像を把握できたので、クイック スタートのいずれかを行います。 プラットフォームについて学習する場合は、Excel クイック スタートをお勧めします。 Visual Studio に基づくバージョンがあります。また、node.js と Visual Studio Code に基づくバージョンがあります。

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js および Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>手順 3: コーディング

オーナーズ マニュアルを読んでも、理解することはできません。この [ Excel チュートリアル](../tutorials/excel-tutorial.md)を使用してコーディングを開始してください。 Office JavaScript ライブラリとアドインのマニフェストにある一部の XML を使用します。 後の手順において、両方の背景がわかりやすくなっているため、何も記憶する必要はありません。

## <a name="step-4-understand-the-javascript-library"></a>手順 4: JavaScript ライブラリを理解する

まず、Microsoft Learn トレーニング: Office JavaScript API について理解するチュートリアルを使用して [、Office JavaScript ライブラリの全体像を確認します](/training/modules/understand-office-javascript-apis/index)。

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
  - [Microsoft 365 開発者プログラムについて](https://developer.microsoft.com/microsoft-365/dev-program)
