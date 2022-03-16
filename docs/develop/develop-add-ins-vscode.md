---
title: Visual Studio Code を使用して Office アドインを開発する
description: Visual Studio Code を使用して Office アドインを開発する方法について説明します。
ms.date: 02/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 6710884a9bc751e6a94607581223dabaea0bce3b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63511288"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>Visual Studio Code を使用して Office アドインを開発する

この記事では、[Visual Studio Code (VS Code)](https://code.visualstudio.com) を使用して Office アドインを開発する方法について説明します。

> [!NOTE]
> Visual Studio を使用して Office アドインを作成する方法については、「[Visual Studio を使用して Office アドインを作成する](develop-add-ins-visual-studio.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>Yeoman ジェネレーターを使用してアドイン プロジェクトを作成する

統合開発環境 (IDE) として VS Code を使用している場合、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で Office アドイン プロジェクトを作成する必要があります。Yeoman ジェネレーターは、VS Code またはその他のエディターで管理できる Node.js プロジェクトを作成します。

Yeoman ジェネレーターを使用して Office アドインを作成するには、作成するアドインの種類に対応する [5 分間のクイック スタート](../index.yml)の指示に従います。

## <a name="develop-the-add-in-using-vs-code"></a>VS Code を使用してアドインを開発する

Yeoman ジェネレーターがアドイン プロジェクトの作成を完了したら、VS Code でプロジェクトのルート フォルダーを開きます。

[!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

Yeoman ジェネレーターは、機能が制限された基本的なアドインを作成します。 VS Code で[マニフェスト](add-in-manifests.md)、HTML、JavaScript または TypeScript、および CSS ファイルを編集することにより、アドインをカスタマイズできます。 Yeoman ジェネレーターが作成するアドイン プロジェクトのプロジェクト構造とファイルの概要については、作成したアドインの種類に対応する [5 分間のクイック スタート](../index.yml)内の Yeoman ジェネレーターのガイダンスを参照してください。

## <a name="test-and-debug-the-add-in"></a>アドインのテストとデバッグ

Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。 詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。

## <a name="publish-the-add-in"></a>アドインを発行する

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>関連項目

- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインの公開](../publish/publish.md)