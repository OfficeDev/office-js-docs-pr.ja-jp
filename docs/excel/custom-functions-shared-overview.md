---
ms.date: 02/13/2020
description: カスタム関数、リボン ボタン、作業ウィンドウのコードを同じ JavaScript ランタイムで実行して、さまざまなアドインでシナリオを調整する方法について説明します。
title: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)
localization_priority: Priority
ms.openlocfilehash: d9d73a5ae2ff1da09d1a5fd7d02514cb28be0e2d
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284128"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a>概要: 共有の JavaScript ランタイムでアドイン コードを実行する (プレビュー)

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

Windows または Mac で Excel を実行する場合、アドインは、リボン ボタン、カスタム関数、作業ウィンドウのコードを別の JavaScript ランタイム環境で実行します。 これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。

ただし、Excel アドインを構成すれば、同じ JavaScript ランタイム (共有ランタイムとも呼ばれる) でコードを共有できるようになります。 これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。

共有ランタイムを構成すると、次のシナリオが可能になります。

- アドインに、リボン、作業ウィンドウ、カスタム関数のすべてがアクセスできる共有の DOM が含まれます。
- カスタム関数で CORS がすべてサポートされます。
- カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。
- ドキュメントを開いてすぐに、アドインでコードを実行できます。
- 作業ウィンドウが閉じられた後でも、アドインでコードの実行を続けられます。

共有ランタイムを使用して作業ウィンドウでカスタム関数を実行すると、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、別のプラットフォームのブラウザー インスタンスで実行されます。また、Excel アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。 次の図は、カスタム関数、リボン UI、作業ウィンドウのコードがすべて同じ JavaScript ランタイム内で実行される様子を示しています。

![Excel でカスタム関数をリボン ボタンと作業ウィンドウと一緒に共有ランタイムで実行](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a>共有ランタイムでカスタム関数を実行するときの違い

Excel アドイン プロジェクトを構成して、共有ランタイムでカスタム関数を実行する場合、カスタム関数のランタイムを使用するのとは異なる点がいくつかあります。

### <a name="storage"></a>ストレージ

作業ウィンドウ、カスタム関数、リボン UI の間でデータを共有するための**ストレージ** API を使用する必要がなくなりました。 **ウィンドウ** オブジェクトにグローバル変数を入力するか、お好みの状態管理アプローチを使うことができます。

### <a name="authentication"></a>認証

認証の一環としてトークンを受け取る場合、作業ウィンドウ、カスタム関数、リボン UI 間でそのトークンを共有するために **ストレージ** API を使用する必要はありません。 お好みのストレージ方法で `localStorage` などの保存場所で共有することができます。

### <a name="dialog-api"></a>ダイアログ API

**OfficeRuntime.Dialog** API を使ってカスタム関数からのダイアログを表示する必要はなくなります。 カスタム関数、リボン ボタン、作業ウィンドウに対して、同じ[ダイアログ API](../develop/dialog-api-in-office-add-ins.md) を使うことができます。

### <a name="debugging"></a>デバッグ

共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。 開発者ツールを使用する必要があります。 さらに詳しい情報については、「[Windows 10 で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)」を参照してください。

## <a name="get-started"></a>使用を開始する

共有ランタイムでカスタム関数を実行するように Excel のアドイン プロジェクトを構成する方法については、「[共有の JavaScript ランタイムを使用するように Excel アドインを構成する (プレビュー)](configure-your-add-in-to-use-a-shared-runtime.md)」を参照してください。

## <a name="give-us-feedback"></a>ご意見をお寄せください

この機能について、ご意見をお待ちしております。 バグや問題が発生したり、この機能について要求がございましたら、[office-js repo](https://github.com/OfficeDev/office-js) で GitHub に関する問題を作成してお知らせください。

## <a name="see-also"></a>関連項目

共有ランタイムの関連記事の一覧
- [チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する (プレビュー)](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [カスタム関数から Excel API を呼び出す (プレビュー)](call-excel-apis-from-custom-function.md)