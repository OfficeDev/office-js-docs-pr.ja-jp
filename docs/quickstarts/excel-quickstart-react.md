---
title: React を使用して Excel 作業ウィンドウ アドインを構築する
description: Office JS API と React を使用して単純な Excel 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 08/04/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 11f31d4145b83ad1efbef441cd78c4ad62410df3f371a991f9f7d94cd75db899
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097517"
---
# <a name="use-react-to-build-an-excel-task-pane-add-in"></a>React を使用して Excel 作業ウィンドウ アドインを構築する

この記事では、React と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。

## <a name="prerequisites"></a>前提条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using React framework`
- **Choose a script type: (スクリプトの種類を選択)** `TypeScript`
- **What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`
- **Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Excel`

![プロジェクトの種類が React フレームワークに設定されている Yeoman Office アドイン ジェネレーター コマンドライン インターフェイスのスクリーンショット。](../images/yo-office-excel-react-2.png)

ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>プロジェクトを確認する

Yeoman ジェネレーターで作成したアドイン プロジェクトには、基本的なアドインの作業ウィンドウのサンプル コードが含まれています。 アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。 アドインを試す準備ができたら、次のセクションに進みます。

- プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。 **manifest.xml** ファイルの詳細については、「[Office アドインの XML マニフェスト](../develop/add-in-manifests.md)」を参照してください。
- **./src/taskpane/taskpane.html** ファイルは作業ウィンドウの HTML フレームワークを定義し、**./src/taskpane/components** フォルダー内のファイルは作業ウィンドウ UI のさまざまな部分を定義します。
- **./src/taskpane/taskpane.css** ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。
- **./src/taskpane/components/App.tsx** ファイルには、作業ウィンドウと Excel の間のやり取りを容易にする Office JavaScript API コードが含まれています。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート フォルダーに移動します。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

1. Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。

    ![[作業ウィンドウの表示] ボタンが強調表示されている Excel ホーム メニューのスクリーンショット。](../images/excel-quickstart-addin-3b.png)

1. ワークシート内で任意のセルの範囲を選択します。

1. 作業ウィンドウの下部で、**[実行]** リンクを選択して、選択範囲の色を黄色に設定します。

    ![アドイン作業ウィンドウが開いており、アドイン作業ウィンドウで [実行] ボタンが強調表示されている Excel のスクリーンショット。](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a>次の手順

おめでとうございます! これで React を使用して Excel 作業ウィンドウ アドインを作成できました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。

> [!div class="nextstepaction"]
> [Excel アドインのチュートリアル](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>関連項目

- [Excel アドインのチュートリアル](../tutorials/excel-tutorial.md)
- [Office アドインの Excel JavaScript オブジェクト モデル](../excel/excel-add-ins-core-concepts.md)
- [Excel アドインのコード サンプル](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Excel JavaScript API リファレンス](../reference/overview/excel-add-ins-reference-overview.md)