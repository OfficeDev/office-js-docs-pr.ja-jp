---
title: Visual Studio Code を使用して Office アドインを開発する
description: Visual Studio Code を使用して Office アドインを開発する方法について説明します。
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 7bd12f6d6e4aff6e8a80f9e9c2e5042b726eca0c
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741100"
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

Yeoman ジェネレーターを使用して Office アドインを作成するには、作成するアドインの種類に対応する [5 分間のクイック スタート](/office/dev/add-ins/)の指示に従います。

## <a name="develop-the-add-in-using-vs-code"></a>VS Code を使用してアドインを開発する

Yeoman ジェネレーターがアドイン プロジェクトの作成を完了したら、VS Code でプロジェクトのルート フォルダーを開きます。 

> [!TIP]
> Windows では、コマンド ラインからプロジェクトのルート ディレクトリに移動し、`code .` を入力して VS Code でそのフォルダーを開くことができます。 Mac では、VS Code でプロジェクト フォルダーを開くためにそのコマンドを使用する前に、[`code` コマンドをパスに追加する](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line)必要があります。

Yeoman ジェネレーターは、機能が制限された基本的なアドインを作成します。 VS Code で[マニフェスト](add-in-manifests.md)、HTML、JavaScript または TypeScript、および CSS ファイルを編集することにより、アドインをカスタマイズできます。 Yeoman ジェネレーターが作成するアドイン プロジェクトのプロジェクト構造とファイルの概要については、作成したアドインの種類に対応する [5 分間のクイック スタート](/office/dev/add-ins/)内の Yeoman ジェネレーターのガイダンスを参照してください。

## <a name="test-and-debug-the-add-in"></a>アドインのテストとデバッグ

Office アドインのテスト、デバッグ、およびトラブルシューティングの方法は、プラットフォームによって異なります。 詳細については、「[Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)」を参照してください。

## <a name="publish-the-add-in"></a>アドインを発行する

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>関連項目

- [Office アドインの中心概念](../overview/core-concepts-office-add-ins.md)
- [Office アドインを開発する](../develop/develop-overview.md)
- [Office アドインを設計する](../design/add-in-design.md)
- [Office アドインのテストとデバッグ](../testing/test-debug-office-add-ins.md)
- [Office アドインを発行する](../publish/publish.md)
