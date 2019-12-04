---
title: Visual Studio Code を使用して Office アドインを開発する
description: Visual Studio Code を使用して Office アドインを開発する方法
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: a18d8a74ff269b32e83c836b06629850873e507b
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670495"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>Visual Studio Code を使用して Office アドインを開発する

この記事では、[Visual Studio Code (VS Code)](https://code.visualstudio.com) を使用して Office アドインを開発する方法について説明します。

> [!NOTE]
> Visual Studio を使用して Office アドインを作成する方法については、「[Visual Studio での Office アドインの作成とデバッグ](create-and-debug-office-add-ins-in-visual-studio.md)」を参照してください。

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>Yeoman ジェネレーターを使用してアドイン プロジェクトを作成する

統合開発環境 (IDE) として VS Code を使用している場合、[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)で Office アドイン プロジェクトを作成する必要があります。Yeoman ジェネレーターは、VS Code またはその他のエディターで管理できる Node.js プロジェクトを作成します。 

Yeoman ジェネレーターを使用して Office アドインを作成するには、作成するアドインの種類に対応する [5 分間のクイック スタート](../index.md)の指示に従います。

## <a name="develop-the-add-in-using-vs-code"></a>VS Code を使用してアドインを開発する

Yeoman ジェネレーターがアドイン プロジェクトの作成を完了したら、VS Code でプロジェクトのルート フォルダーを開きます。 

> [!TIP]
> Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。 Mac では、VS Code でプロジェクト フォルダーを開くためにそのコマンドを使用する前に、[`code` コマンドをパスに追加する](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)必要があります。

Yeoman ジェネレーターは、機能が制限された基本的なアドインを作成します。 VS Code で[マニフェスト](add-in-manifests.md)、HTML、JavaScript または TypeScript、および CSS ファイルを編集することにより、アドインをカスタマイズできます。 Yeoman ジェネレーターが作成するアドイン プロジェクトのプロジェクト構造とファイルの概要については、作成したアドインの種類に対応する [5 分間のクイック スタート](../index.md)内の Yeoman ジェネレーターのガイダンスを参照してください。

## <a name="test-and-debug-the-add-in"></a>アドインのテストとデバッグ

Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。 詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。

## <a name="publish-the-add-in"></a>アドインを発行する

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>関連項目

- [5 分間のクイック スタート](../index.md)
- [Script Lab を使用して Office JavaScript API を探索する](../overview/explore-with-script-lab.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインを展開し、発行する](../publish/publish.md)