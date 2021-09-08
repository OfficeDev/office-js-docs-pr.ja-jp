---
title: 作業ウィンドウからデバッガーをアタッチする
description: 作業ウィンドウからデバッガーをアタッチする方法について学習する
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 930c5db7a9e1a0499a759d70e0bc94e64c5b7ae5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938354"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>作業ウィンドウからデバッガーをアタッチする

Windows での Office 2016 のビルド 77xx.xxxx 以降では、作業ウィンドウからデバッガーをアタッチすることができます。デバッガーのアタッチ機能によって、デバッガーが適切な Internet Explorer プロセスに直接アタッチされます。デバッガーは、Yeoman Generator、Visual Studio Code、Node.js、Angular、その他のツールのどれを使用しているかに関係なくアタッチすることができます。

**デバッガーのアタッチ** ツールを起動するのには、作業ウィンドウの右上隅を選択して **パーソナリティ** メニューを有効にします (以下の図の赤い円で示す通り)。

> [!NOTE]
> - 現在、サポートされているデバッガー ツールは、更新プログラム[3](/previous-versions/mt752379(v=vs.140))以降Visual Studio [2015](https://www.visualstudio.com/downloads/)でサポートされています。 インストールされていない場合Visual Studioデバッガーの接続オプションを選択しても、操作は実行されません。
> - **[デバッガーのアタッチ]** ツールでデバッグできるのは、クライアント側の JavaScript だけです。 Node.js サーバーなど、サーバー側のコードをデバッグするには、多くのオプションがあります。 Visual Studio Code でデバッグするための詳しい方法については、「[VS Code で Node.js をデバッグする](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)」を参照してください。 Visual Studio Code を使用していない場合は、「Node.js のデバッグ」または「{サーバー名} のデバッグ」で検索してください。

![[デバッガーのアタッチ] メニューのスクリーンショット。](../images/attach-debugger.png)

**デバッガーのアタッチ** を選択するこれにより、次の図のように、**Visual Studio Just-in-Time デバッガー** ダイアログ ボックスが起動します。

![[JIT デバッガー Visual Studioのスクリーンショットです。](../images/visual-studio-debugger.png)

Visual Studio では、**ソリューション エクスプローラー** 内にコード ファイルが表示されます。   Visual Studio でデバッグするコードの行にブレークポイントを設定することができます。

> [!NOTE]
> [パーソナリティ] メニューが表示されない場合は、Visual Studio を使用してアドインをデバッグできます。 作業ウィンドウ アドインが [作業ウィンドウ] ウィンドウで開いているOffice、次の手順を実行します。
>
> 1. Visual Studio で、**[デバッグ]** > **[プロセスにアタッチ]** の順に選択します。
> 2. **使用可能なプロセス** で、[アドインが Internet Explorer または Microsoft Edge のどちらを使用しているか](../concepts/browsers-used-by-office-web-add-ins.md)に応じて、使用可能なすべての `Iexplore.exe` プロセス *または* 使用可能なすべての `MicrosoftEdge*.exe` プロセスの *どちらか* を選択し、[**添付**] ボタンを選択します。

Visual Studio でのデバッグの詳細については、以下を参照してください。

- DOM Explorer を Visual Studio で起動して使用するには、ブログ記事「[新しいプロジェクト テンプレートを使って見栄えの良い Office 用アプリをビルドする](/archive/blogs/officeapps/building-great-looking-apps-for-office-using-the-new-project-templates)」の[ヒントとコツ](/archive/blogs/officeapps/building-great-looking-apps-for-office-using-the-new-project-templates#tips_tricks) セクションのヒント 4 を参照してください。
- ブレークポイントの設定については、「[ブレークポイントの使用](/visualstudio/debugger/using-breakpoints?view=vs-2015&preserve-view=true)」を参照してください。
- F12 を使用するには、「[F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))」を参照してください。
- Microsoft Edge 開発者ツールを使用するには、「[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)」を参照してください。

## <a name="see-also"></a>関連項目

- [Visual Studio で Office アドインをデバッグする](../develop/debug-office-add-ins-in-visual-studio.md)
- [Office アドインを発行する](../publish/publish.md)
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)