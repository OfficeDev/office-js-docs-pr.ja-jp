---
title: Visual Studio を使用して Office アドインを開発する
description: Visual Studio を使用して Office アドインを開発する方法
ms.date: 07/08/2021
localization_priority: Priority
ms.openlocfilehash: bc837c0fe399cc6669cb0efcf2531e438f922dfe
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773791"
---
# <a name="develop-office-add-ins-with-visual-studio"></a>Visual Studio を使用して Office アドインを開発する

この記事では、Visual Studio を使用して Office アドインを開発する方法について説明します。 アドインの作成が既に完了している場合は、「[Visual Studio を使用して アドインを開発する](#develop-the-add-in-using-visual-studio)」セクションに進んでください。

> [!NOTE]
> Visual Studio を使用する代わりに、Office アドイン用の Yeoman ジェネレーターと VS コードを使用して Office アドインを作成することもできます。 この選択肢の詳細については、「[Office アドインの作成 ](../develop/develop-overview.md)」#creating-an-office-add-in)を参照してください。

## <a name="create-the-add-in-project-using-visual-studio"></a>Visual Studio を使用してアドイン プロジェクトを作成する

Visual Studio は、Excel、Outlook、Word、および PowerPoint 用の Office アドインの作成に使用できます。 Office アドイン プロジェクトは Visual Studio ソリューションの一部として作成され、HTML、CSS、および JavaScript が使用されます。 Visual Studio を使用して Office アドインを作成するには、作成するアドインに対応するクイック スタートの指示に従います。

- [Excel クイック スタート](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Outlook クイック スタート](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [Word クイック スタート](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [PowerPoint クイック スタート](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

Visual Studio では、OneNote または Project 用の Office アドインの作成はサポートされていません。 これらのいずれかのアプリケーション用の Office アドインを作成するには、[OneNote クイック スタート](../quickstarts/onenote-quickstart.md) または [Project クイック スタート](../quickstarts/project-quickstart.md) で説明するように、Office アドイン用の Yeoman ジェネレーターを使用する必要があります。

## <a name="develop-the-add-in-using-visual-studio"></a>Visual Studio を使用してアドインを開発する

Visual Studio では、機能が制限された基本的なアドインが作成されます。 [マニフェスト](add-in-manifests.md)、HTML、JavaScript、および CSS の各ファイルを Visual Studio で編集することで、アドインをカスタマイズできます。 Visual Studio により作成されるアドイン プロジェクトのプロジェクト構造とファイルの概要については、アドインを作成するために実行したクイック スタート内の Visual Studio ガイダンスを参照してください。

> [!TIP]
> Office アドインは Web アプリケーションであるため、アドインをカスタマイズするには、少なくとも Web 開発の基本的なスキルが必要です。 JavaScript を使い慣れていない場合は、[Mozilla の JavaScript チュートリアル](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)をご覧になることをお勧めします。

アドインをカスタマイズするには、このドキュメントの [「中心概念」 > 「開発」](develop-overview.md) 項目で説明する概念の他、作成するるアドインに対応するドキュメント内の、アプリケーション固有の項目 (例: [Excel](../excel/index.yml)) で説明する概念を理解する必要があります。

## <a name="test-and-debug-the-add-in"></a>アドインのテストとデバッグ

Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。 詳細については、「[Visual Studio で Office アドインをデバッグする](debug-office-add-ins-in-visual-studio.md)」および「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。

## <a name="publish-the-add-in"></a>アドインを発行する

Office アドインは、Web アプリケーションとマニフェスト ファイルで構成されます。Web アプリケーションはアドインのユーザー インターフェイスと機能を定義しますが、マニフェストは Web アプリケーションの場所を指定し、アドインの設定と機能を定義します。

Visual Studio で開発中のアドインは、ローカル Web サーバー上 (`localhost`) で実行されます。 アドインが正常に機能し、他のユーザーがアクセスできるように公開する準備ができた場合、次の手順を実行する必要があります。

1. Web アプリケーションを Web サーバーまたは Web ホスティング サービス (例: Microsoft Azure) に展開します。
2. マニフェストを更新して、展開されたアプリケーションの URL を指定します。
3. [Office アドインを展開](../publish/publish.md)するために使用する方法を選択し、指示に従ってマニフェスト ファイルを公開します。

## <a name="see-also"></a>関連項目

- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)
