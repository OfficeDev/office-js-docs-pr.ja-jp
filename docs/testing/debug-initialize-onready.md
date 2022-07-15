---
title: initialize メソッドと onReady メソッドをデバッグする
description: Office.initialize メソッドと Office.onReady メソッドをデバッグする方法について説明します。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed6e69a52f3f4534db075daf62c171d4806e89d4
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797717"
---
# <a name="debug-the-initialize-and-onready-methods"></a>initialize メソッドと onReady メソッドをデバッグする

> [!NOTE]
> この記事では、 [Office アドインの初期化について](../develop/initialize-add-in.md)理解していることを前提としています。

[Office.initialize](/javascript/api/office#office-office-initialize-function(1)) メソッドと [Office.onReady](/javascript/api/office#office-office-onready-function(1)) メソッドのデバッグのパラドックスは、デバッガーが実行されているプロセスにのみアタッチできるということですが、これらのメソッドは、アドインのランタイム プロセスが起動すると、デバッガーがアタッチする前にすぐに実行されます。 ほとんどの場合、デバッガーがアタッチされた後にアドインを再起動しても、アドインを再起動しても元のランタイム プロセス *とアタッチされたデバッガー* が閉じられ、デバッガーがアタッチされていない新しいプロセスが開始されるため、役に立ちません。

幸いにも、例外があります。 これらのメソッドは、次の手順でOffice on the webを使用してデバッグできます。

1. アドインをサイドロードし、Office on the webで実行します。 これは通常、アドインの作業ウィンドウを開くか、 [関数コマンド](../design/add-in-commands.md#types-of-add-in-commands)を実行することによって行われます。 *アドインは、デスクトップ Office とは別のプロセスではなく、ブラウザー プロセス全体で実行されます。*
1. ブラウザーの開発者ツールを開きます。 これは通常、F12 キーを押すことによって行われます。 ツールのデバッガーは、ブラウザー プロセスにアタッチされます。
1. 必要に応じて、または`Office.onReady`メソッド内のコードにブレークポイントを`Office.initialize`適用します。
1. 手順 1 で行ったように、*アドインの作業ウィンドウまたは関数コマンドを再起動* します。 このアクションは、ブラウザー プロセスまたはデバッガーを閉じ *ません* 。 または`Office.onReady`メソッドが`Office.initialize`再度実行され、ブレークポイントで処理が停止します。

> [!TIP]
> 詳細については、「[Office on the webでのアドインのデバッグ](debug-add-ins-in-office-online.md)」を参照してください。 
