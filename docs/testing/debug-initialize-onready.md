---
title: initialize 関数と onReady 関数をデバッグする
description: Office.initialize 関数と Office.onReady 関数をデバッグする方法について説明します。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d26c0f2686477f74fbe65550b7d35cc69e7be71
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958847"
---
# <a name="debug-the-initialize-and-onready-functions"></a>initialize 関数と onReady 関数をデバッグする

> [!NOTE]
> この記事では、 [Office アドインの初期化について](../develop/initialize-add-in.md)理解していることを前提としています。

[Office.initialize](/javascript/api/office#office-office-initialize-function(1)) 関数と [Office.onReady](/javascript/api/office#office-office-onready-function(1)) 関数のデバッグのパラドックスは、デバッガーが実行されているプロセスにのみアタッチできるということですが、これらの関数は、アドインのランタイム プロセスが起動すると、デバッガーがアタッチする前にすぐに実行されます。 ほとんどの場合、デバッガーがアタッチされた後にアドインを再起動しても、アドインを再起動しても元のランタイム プロセス *とアタッチされたデバッガー* が閉じられ、デバッガーがアタッチされていない新しいプロセスが開始されるため、役に立ちません。

幸いにも、例外があります。 これらの関数は、次の手順でOffice on the webを使用してデバッグできます。

1. アドインをサイドロードし、Office on the webで実行します。 これは通常、アドインの作業ウィンドウを開くか、 [関数コマンド](../design/add-in-commands.md#types-of-add-in-commands)を実行することによって行われます。 *アドインは、デスクトップ Office とは別のプロセスではなく、ブラウザー プロセス全体で実行されます。*
1. ブラウザーの開発者ツールを開きます。 これは通常、F12 キーを押すことによって行われます。 ツールのデバッガーは、ブラウザー プロセスにアタッチされます。
1. 必要に応じて、または`Office.onReady`関数内のコードにブレークポイントを`Office.initialize`適用します。
1. 手順 1 で行ったように、*アドインの作業ウィンドウまたは関数コマンドを再起動* します。 このアクションは、ブラウザー プロセスまたはデバッガーを閉じ *ません* 。 または`Office.onReady`関数が`Office.initialize`再度実行され、ブレークポイントで処理が停止します。

> [!TIP]
> 詳細については、「[Office on the webでのアドインのデバッグ](debug-add-ins-in-office-online.md)」を参照してください。
