---
title: VSTO アドイン開発者向けガイド
description: 熟練した VSTO アドイン開発者に向けた Office Web アドイン学習リソースへの推奨パス。
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 1dca15a4d286e3bfa5b7ba4a502bb9161bf3257f
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741065"
---
# <a name="vsto-add-in-developers-guide"></a>VSTO アドイン開発者向けガイド

Windows で動作する Office アプリケーション用の VSTO アドインを作成しました。そしてここからは、Office を Windows、Mac、オンライン バージョンの Office スイートで動作するように拡張するための新しい方法である Office Web アドインについて説明します。

Office Web アドインのオブジェクト モデルは Excel、Word、その他の Office アプリケーションのオブジェクト モデルと似たようなパターンをたどるので、それらのオブジェクト モデルへの理解が大きな助けとなるでしょう。 ただし、いくつか課題があります。

- C# や Visual Basic .NET ではなく、別の言語 (JavaScript または TypeScript のいずれか) を使用して作業することになります。 (後述するように、既存のコードの一部を Web アドインで再利用する方法もあります)。
- Office Web アドインは、VSTO アドインとは別に展開されます。
- Office Web アドインは、Office アプリケーションに組み込まれた簡易ブラウザー ウィンドウで動作する Web アプリケーションなので、Web アプリケーションの基本的な理解と、それらがどのように Web サーバーやクラウド アカウントでホストされるかについてを理解しておく必要があります。 

これらの理由により、この記事の多くは、Office 拡張機能の完全な初心者向けの学習パスである、[初心者向けガイド](learning-path-beginner.md) と重複しています。 VSTO アドイン開発者が経験を活用し、既存のコードを再利用できるようにするための追加の学習リソースを追加しました。

## <a name="step-0-prerequisites"></a>手順 0: 前提条件

- Office Web アドイン (Office アドインとも呼ばれる) は、Office に組み込まれている基本 Web アプリケーションです。 まず、Web アプリケーションの基本について説明し、次に、Web でのホスト方法について説明します。 インターネット、書籍、オンライン コース上にこれについての膨大な情報があります。 Web アプリケーションに関する事前知識がまったくない場合、Bing で "Web アプリとは?" と検索することから始めることを お勧めします。
- Office アドインの作成に使用する主要なプログラミング言語は、JavaScript または TypeScript です。 TypeScript は、JavaScript の厳密に型指定されたバージョンと考えることができます。 これらの言語のいずれにも慣れておらず、VBA、VB.Net、C# の経験がある場合、TypeScript の方が学習しやすいかもしれません。 繰り返しになりますが、インターネット、書籍、オンライン コース上に、これらの言語に関する豊富な情報があります。

## <a name="step-1-begin-with-fundamentals"></a>手順 1: 基本から始める

今にもコーディングを始めたいと考えておられるかもしれませんが、IDE やコード エディターを開く前に、Office アドインについて、以下をお読みください。

- [Office アドイン プラットフォームの概要](office-add-ins.md): Office Web アドインとは何であるか、VSTO アドインなどの Office を拡張する以前の方法との違いを説明します。
- [Office アドインの開発](../develop/develop-overview.md): ツール、アドイン UI の作成、JavaScript API を使用する Office ドキュメントの操作を含む、Office アドインの開発とライフサイクルの概要を説明します。

これらの記事には多くのリンクが含まれていますが、Office アドインに移行している場合は、これらを読み、次のセクションに進むときに、ここに戻ることをお勧めします。

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>手順 2: ツールをインストールし、最初のアドインを作成する

全体像を把握できたので、クイック スタートのいずれかを行います。 プラットフォームについて学習する場合は、Excel クイック スタートをお勧めします。 Visual Studio をベースにしたバージョンと、Node.js と Visual Studio Code をベースにしたバージョンがあります。 VSTO アドインから移行している場合は、Visual Studio バージョンの方が作業がしやすいかもしれません。

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js および Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>手順 3: コーディング

オーナーズ マニュアルを読んでも、理解することはできません。この [ Excel チュートリアル](../tutorials/excel-tutorial.md)を使用してコーディングを開始してください。 Office JavaScript ライブラリとアドインのマニフェストにある一部の XML を使用します。 後の手順で両方の背景を学習することになるので、何も覚える必要はありません。

## <a name="step-4-understand-the-javascript-library"></a>手順 4: JavaScript ライブラリを理解する

Microsoft Learn「[Office JavaScript API について理解する](/learn/modules/intro-office-add-ins/3-apis)」のこのチュートリアルで、Office JavaScript ライブラリの全体像を把握します。

次に、API を実行して調査するためのサンドボックスである [Script Lab ツール](explore-with-script-lab.md)を使用して、Office JavaScript API を学習します。

### <a name="special-resource-for-vsto-add-in-developers"></a>VSTO アドイン開発者向けの特別なリソース

サンプルのアドインを見るには、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) が良いでしょう。 これは VSTO アドインと Office Web アドインの共通点や違いを強調するために作成されたもので、このサンプルの readme では比較する上での重要なポイントについてを紹介しています。

## <a name="step-5-understand-the-manifest"></a>手順 5: マニフェストを理解する

Web アドイン マニフェストの目的を理解し、[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)で XML マークアップの概要について説明します。

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a>手順 6 (VSTO 開発者のみ): VSTO コードを再利用する

VSTO アドインのコードをサーバー上の Web アプリケーションのバックエンドへと移動し、JavaScript や TypeScript で Web API として利用できるようにすることにより、Office Web アドインで VSTO アドインのコードを再利用することができます。 ガイダンスについては、「[チュートリアル: 共有コード ライブラリを使用して VSTO アドインと Office アドインの間でコードを共有する](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md)」を参照してください。

## <a name="next-steps"></a>次の手順

おめでとうございます。Office Web アドインの VSTO アドイン開発者向け学習パスを完了しました! ドキュメントの詳細については、以下をご覧ください。

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
  - [Resources](../resources/resources-links-help.md)
  - [Microsoft 365 開発者プログラムについて学ぶ](https://developer.microsoft.com/microsoft-365/dev-program)
