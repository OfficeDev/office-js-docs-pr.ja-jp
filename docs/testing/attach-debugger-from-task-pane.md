---
title: 作業ウィンドウからデバッガーをアタッチする
description: 作業ウィンドウからデバッガーをアタッチする方法について学習します。
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0363b7966ab3da11167cb4b0cd324df28fd9efb3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744752"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>作業ウィンドウからデバッガーをアタッチする

一部の環境では、デバッガーを既に実行Officeアドインに接続できます。 これは、既にステージングまたは運用中のアドインをデバッグする場合に役立ちます。 アドインを開発およびテストする場合は、「アドインのデバッグのOffice[」を参照してください](debug-add-ins-overview.md)。

この記事で説明する手法は、次の条件が満たされている場合にのみ使用できます。

- アドインは、アプリのOfficeでWindows。
- コンピューターは、エッジ (Windows ベースの) webview コントロール WebView2 をOfficeバージョンと Chromiumバージョンの組み合わせを使用しています。 使用しているブラウザーを確認するには、「アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

デバッガーを起動するには、作業ウィンドウの右上隅を選択して [パーソナリティ] メニューをアクティブにします (次の図の赤い円に示すように)。

![[デバッガーのアタッチ] メニューのスクリーンショット。](../images/attach-debugger.png)

[デバッガー **の接続] を選択します**。 これにより、Microsoft Edge (Chromium) 開発者ツールが起動します。 「開発ツールを使用してアドインをデバッグする(Microsoft Edgeベース)Chromium[ツールを使用します](debug-add-ins-using-devtools-edge-chromium.md)。

## <a name="see-also"></a>関連項目

- [Office アドインのデバッグの概要](debug-add-ins-overview.md)
