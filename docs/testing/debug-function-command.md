---
title: 非共有ランタイムを使用する関数コマンドをデバッグする
description: 関数コマンドをデバッグする方法について説明します。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d2be148c05f88837610b8563c2e61618d1c37775
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423203"
---
# <a name="debug-a-function-command-with-a-non-shared-runtime"></a>非共有ランタイムを使用する関数コマンドをデバッグする

> [!IMPORTANT]
> [アドインが共有ランタイムを使用するように構成されている](../develop/configure-your-add-in-to-use-a-shared-runtime.md)場合は、作業ウィンドウの背後にあるコードと同様に、関数コマンドの背後にあるコードをデバッグします。 [Office アドインのデバッグに](debug-add-ins-overview.md)関する記事を参照し、[共有ランタイム](runtimes.md#shared-runtime)を使用するアドインの関数コマンドは、その記事で説明されているように特別なケース *ではありません*。 

> [!NOTE]
> この記事では、 [関数コマンド](../design/add-in-commands.md#types-of-add-in-commands)について理解していることを前提としています。

関数コマンドには UI がないため、デバッガーをデスクトップ Office で実行するプロセスにアタッチできません。 (Windows で開発されている Outlook アドインは、これに対する例外です。 この記事[の後半の「Windows 上の Outlook アドインで関数コマンドをデバッグ](#debug-function-commands-in-outlook-add-ins-on-windows)する」を参照してください)。そのため、非共有ランタイムを使用するアドインの関数コマンドは、関数がブラウザー プロセス全体で実行されるOffice on the webでデバッグする必要があります。 次の手順を使用します。

1. Office on the webアドインをサイドロードし、関数コマンドを実行するボタンまたはメニュー項目を選択します。 これは、関数コマンドのコード ファイルを読み込む際に必要です。 
1. ブラウザーの開発者ツールを開きます。 これは通常、F12 キーを押すことによって行われます。 ツールのデバッガーは、ブラウザー プロセスにアタッチされます。
1. 関数コマンドに必要に応じて、コードにブレークポイントを適用します。
1. 関数コマンドを再実行します。 プロセスはブレークポイントで停止します。 

> [!TIP]
> 詳細については、「[Office on the webでのアドインのデバッグ](debug-add-ins-in-office-online.md)」を参照してください。

## <a name="debug-function-commands-in-outlook-add-ins-on-windows"></a>Windows 上の Outlook アドインで関数コマンドをデバッグする

開発用コンピューターが Windows の場合は、Outlook デスクトップで関数コマンドをデバッグする方法があります。 [Outlook アドインのデバッグ関数コマンドに関する](../outlook/debug-ui-less.md)説明を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインのランタイム](runtimes.md)
