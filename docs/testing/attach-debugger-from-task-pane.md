---
title: 作業ウィンドウからデバッガーをアタッチする
description: 作業ウィンドウからデバッガーをアタッチする方法について学習する
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 42f987dc4d19ad17140316d82634acf8695fd88d
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263073"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>作業ウィンドウからデバッガーをアタッチする

一部の環境では、デバッガーを既に実行Officeアドインに接続できます。 これは、既にステージングまたは運用中のアドインをデバッグする場合に役立ちます。 アドインを開発およびテストする場合は、「アドインのデバッグの概要」をOffice[参照してください](debug-add-ins-overview.md)。

この記事で説明する手法は、次の条件が満たされている場合にのみ使用できます。

- アドインは、アプリのOfficeでWindows。
- コンピューターは、エッジ (Windowsベース) webview Office WebView2 を使用するバージョンと Chromiumバージョンの組み合わせを使用しています。 使用しているブラウザーを確認するには、「アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

デバッガーを起動するには、作業ウィンドウの右上隅を選択して [パーソナリティ] メニューをアクティブにします (次の図の赤い円に示すように)。

![[デバッガーのアタッチ] メニューのスクリーンショット。](../images/attach-debugger.png)

[デバッガー **の接続] を選択します**。 これにより、Microsoft Edge (Chromium) 開発者ツールが起動します。 「開発ツールを使用してアドインをデバッグする(Microsoft Edgeベース)Chromium[ツールを使用します](debug-add-ins-using-devtools-edge-chromium.md)。

## <a name="see-also"></a>関連項目

- [Office アドインのデバッグの概要](debug-add-ins-overview.md)
