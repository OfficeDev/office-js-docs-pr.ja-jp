---
ms.date: 05/17/2020
description: カスタム関数、リボン ボタン、作業ウィンドウのコードを同じ JavaScript ランタイムで実行して、さまざまなアドインでシナリオを調整する方法について説明します。
title: 共有 JavaScript ランタイムでアドインコードを実行する
localization_priority: Priority
ms.openlocfilehash: afb07c5223e26ba1e1adbf40c7a4b2e4f7c06349
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275932"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a>概要: 共有 JavaScript ランタイムでアドインコードを実行する

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。 これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。

ただし、Excel アドインを構成すれば、同じ JavaScript ランタイム (共有ランタイムとも呼ばれる) でコードを共有できるようになります。 これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。

共有ランタイムを構成すると、次のシナリオが可能になります。

- アドインに、リボン、作業ウィンドウ、カスタム関数のすべてがアクセスできる共有の DOM が含まれます。
- カスタム関数で CORS がすべてサポートされます。
- カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。
- ドキュメントを開いてすぐに、アドインでコードを実行できます。
- 作業ウィンドウが閉じられた後でも、アドインでコードの実行を続けられます。

共有ランタイムを使用して作業ウィンドウでカスタム関数を実行すると、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、別のプラットフォームのブラウザー インスタンスで実行されます。また、Excel アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。 次の図は、カスタム関数、リボン UI、作業ウィンドウのコードがすべて同じ JavaScript ランタイム内で実行される様子を示しています。

![Excel のリボンボタンと作業ウィンドウを使用して共有ランタイムで実行されているカスタム関数](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a>共有ランタイムをセットアップする

共有ランタイムを使用するようにカスタム関数を設定する方法については、「[共有ランタイムの構成](./configure-your-add-in-to-use-a-shared-runtime.md)」の記事を参照してください。

### <a name="debugging"></a>デバッグ

共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。 代わりに、開発者ツールを使用する必要があります。 さらに詳しい情報については、「[Windows 10 で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)」を参照してください。

## <a name="give-us-feedback"></a>ご意見をお寄せください

この機能について、ご意見をお待ちしております。 バグや問題が発生したり、この機能について要求がございましたら、[office-js repo](https://github.com/OfficeDev/office-js) で GitHub に関する問題を作成してお知らせください。

## <a name="see-also"></a>関連項目

- [チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [カスタム関数から Excel Api を呼び出す](call-excel-apis-from-custom-function.md)
