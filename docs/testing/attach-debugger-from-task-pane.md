---
title: 作業ウィンドウからデバッガーをアタッチする
description: 作業ウィンドウからデバッガーを接続する方法について説明します。
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 53cfce211241dbdf3d16e8a126e059a2f2db3f23
ms.sourcegitcommit: b939312ffdeb6e0a0dfe085db7efe0ff143ef873
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/19/2020
ms.locfileid: "44810843"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>作業ウィンドウからデバッガーをアタッチする

In Office 2016 on Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, Node.js, Angular, or another tool. 

**デバッガーのアタッチ** ツールを起動するのには、作業ウィンドウの右上隅を選択して**パーソナリティ** メニューを有効にします (以下の図の赤い円で示す通り)。   

> [!NOTE]
> - 現在サポートされているデバッガー ツールは、[Update 3](https://www.visualstudio.com/downloads/) 以降を適用した [Visual Studio 2015](https://msdn.microsoft.com/library/mt752379.aspx) だけです。 Visual Studio がインストールされていない場合、[**デバッガーのアタッチ**] オプションを選択しても、アクションは実行されません。   
> - You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".

![[デバッガーのアタッチ] メニューのスクリーンショット](../images/attach-debugger.png)

Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image. 

![Visual Studio JIT デバッガー ダイアログのスクリーンショット](../images/visual-studio-debugger.png)

In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.

> [!NOTE]
> [パーソナリティ] メニューが表示されない場合は、Visual Studio を使用してアドインをデバッグできます。 Office で作業ウィンドウ アドインが開いていることを確認してから、次の手順を実行します。
>
> 1. Visual Studio で、**[デバッグ]** > **[プロセスにアタッチ]** の順に選択します。
> 2. **使用可能なプロセス**で、[アドインが Internet Explorer または Microsoft Edge のどちらを使用しているか](../concepts/browsers-used-by-office-web-add-ins.md)に応じて、使用可能なすべての `Iexplore.exe` プロセス*または*使用可能なすべての `MicrosoftEdge*.exe` プロセスの*どちらか*を選択し、[**添付**] ボタンを選択します。

Visual Studio でのデバッグの詳細については、以下を参照してください。

-    DOM Explorer を Visual Studio で起動して使用するには、ブログ記事「[新しいプロジェクト テンプレートを使って見栄えの良い Office 用アプリをビルドする](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates)」の[ヒントとコツ](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) セクションのヒント 4 を参照してください。
-    ブレークポイントの設定については、「[ブレークポイントの使用](/visualstudio/debugger/using-breakpoints?view=vs-2015)」を参照してください。
-    F12 を使用するには、「[F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))」を参照してください。
-   Microsoft Edge 開発者ツールを使用するには、「[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)」を参照してください。

## <a name="see-also"></a>関連項目

- [Visual Studio で Office アドインをデバッグする](../develop/debug-office-add-ins-in-visual-studio.md)
- [Office アドインを発行する](../publish/publish.md)
- [Visual Studio Code の Microsoft Office アドインデバッガーの拡張機能](debug-with-vs-extension.md)