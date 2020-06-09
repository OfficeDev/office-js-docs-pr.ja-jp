---
title: Office ダイアログ API のベスト プラクティスとルール
description: 単一ページアプリケーション (SPA) のベストプラクティスなど、Office ダイアログ API のルールとベストプラクティスを提供します。
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 88c833d91cc16684b5e434d6aff9e77f23bbbdb4
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608272"
---
# <a name="best-practices-and-rules-for-the-office-dialog-api"></a>Office ダイアログ API のベスト プラクティスとルール

この記事では、Office ダイアログ API のルール、ガイドライン、およびベストプラクティスについて説明します。これには、ダイアログの UI を設計し、単一ページアプリケーション (SPA) で API を使用するためのベストプラクティスが含まれています。

> [!NOTE]
> この記事では、「office[アドインで office ダイアログ api を使用](dialog-api-in-office-add-ins.md)する」で説明されている OFFICE ダイアログ api の使用についての基本事項を presupposes しています。
> 
> 「 [Office ダイアログボックスでエラーとイベントを処理する](dialog-handle-errors-events.md)」も参照してください。

## <a name="rules-and-gotchas"></a>ルールと注意事項

- このダイアログボックスは、HTTP ではなく HTTPS の Url にのみ移動できます。
- [Displaydialogasync](/javascript/api/office/office.ui)メソッドに渡される URL は、アドイン自体とまったく同じドメイン内にある必要があります。 サブドメインにすることはできません。 ただし、それに渡されるページは、別のドメイン内のページにリダイレクトすることができます。
- ホストウィンドウは、作業ウィンドウまたはアドインコマンドの UI に含まれない[関数ファイル](../reference/manifest/functionfile.md)の場合がありますが、一度に開くことのできるダイアログボックスは1つだけです。
- ダイアログボックスでは、次の2つの Office Api のみを呼び出すことができます。
  - [Messageparent](/javascript/api/office/office.ui#messageparent-message-)関数。
  - `Office.context.requirements.isSetSupported`(詳細については、「 [Office ホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)」を参照してください)。
- [Messageparent](/javascript/api/office/office.ui#messageparent-message-)関数は、アドイン自体とまったく同じドメイン内のページからのみ呼び出すことができます。

## <a name="best-practices"></a>ベスト プラクティス

### <a name="avoid-overusing-dialog-boxes"></a>過剰な使用ダイアログボックスを回避する

UI 要素を重ねて表示することはお勧めできないため、シナリオで必要な場合を除き、作業ウィンドウでダイアログ ボックスを開かないようにします。 作業ウィンドウの表示領域の使用方法を検討するときには、作業ウィンドウはタブ表示できることに注意してください。 例については、[Excel アドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) のサンプルを参照してください。

### <a name="designing-a-dialog-box-ui"></a>ダイアログボックスの UI を設計する

ダイアログボックスデザインのベストプラクティスについては、「 [Office アドインのダイアログボックス](../design/dialog-boxes.md)」を参照してください。

### <a name="handling-pop-up-blockers-with-office-on-the-web"></a>Office on the web を使用したポップアップ ブロックの処理

Web 上で Office を使用しているときにダイアログボックスを表示しようとすると、ブラウザーのポップアップブロックによってダイアログボックスがブロックされることがあります。 Web 上の Office には、ブラウザーのポップアップブロックに対してアドインのダイアログボックスを例外として使用できる機能があります。 コードによってメソッドが呼び出されると `displayDialogAsync` 、web 上の Office は次のようなプロンプトを開きます。

![ブラウザーでポップアップブロックを発生させないようにするために、アドインが生成できるプロンプト。](../images/dialog-prompt-before-open.png)

ユーザーが [**許可**] を選択すると、[Office] ダイアログボックスが開きます。 ユーザーが [**無視**] を選択すると、プロンプトは閉じられ、[Office] ダイアログボックスは開きません。 代わりに、この `displayDialogAsync` メソッドはエラー12009を返します。 コードでは、このエラーをキャッチして、ダイアログを必要としない別の方法を提供するか、アドインでダイアログを許可する必要があることをユーザーに通知するメッセージを表示する必要があります。 (12009 の詳細については、「 [displayDialogAsync からのエラー](dialog-handle-errors-events.md#errors-from-displaydialogasync)」を参照してください)。

何らかの理由でこの機能をオフにする場合は、コードでオプトアウトする必要があります。この要求は、メソッドに渡される "引数の[選択](/javascript/api/office/office.dialogoptions)" オブジェクトを使用して行われ `displayDialogAsync` ます。 具体的には、オブジェクトにを含める必要があり `promptBeforeOpen: false` ます。 このオプションが false に設定されている場合、web 上の Office では、アドインでダイアログを開くことを許可するかどうかを確認するメッセージは表示されず、[Office] ダイアログボックスは開きません。

### <a name="do-not-use-the-_host_info-value"></a>[ \_ ホスト情報] の値を使用しないでください。 \_

Office は、`_host_info` に渡される URL に `displayDialogAsync` というクエリ パラメーターを自動的に追加します (カスタム クエリ パラメーターが存在する場合は、その後に追加されます。 これはカスタムクエリパラメーターの後に追加されます (存在する場合)。 これは、ダイアログボックスが移動する以降の Url には追加されません。 Microsoft では、この値の内容を変更したり、完全に削除したりすることがあります。この場合、コードでこの値を読み取ることはできません。 ダイアログ ボックスのセッション ストレージには、同じ値が追加されます。 この場合も、*コードではこの値に対する読み取りも書き込みも行わないでください*。

### <a name="best-practices-for-using-the-office-dialog-api-in-an-spa"></a>SPA で Office ダイアログ API を使用するためのベストプラクティス

アドインでクライアント側のルーティングを使用する場合は、通常、単一ページのアプリケーション (SPAs) として、別の HTML ページの URL の代わりに、ルートの URL を[Displaydialogasync](/javascript/api/office/office.ui)メソッドに渡すオプションがあります。 *以下の理由から、そのようにすることをお勧めします。*

> [!NOTE]
> この記事は、エクスプレスベースの web アプリケーションなど、*サーバー側*のルーティングには関連していません。

#### <a name="problems-with-spas-and-the-office-dialog-api"></a>SPAs と Office ダイアログ API の問題

[Office] ダイアログボックスは、JavaScript エンジンの独自のインスタンスを備えた新しいウィンドウに表示されるので、完全な実行コンテキストです。 ルートを渡すと、基本ページとすべての初期化コードとブートストラップコードがこの新しいコンテキストで再び実行され、すべての変数がダイアログボックスの初期値に設定されます。 そのため、この手法では、アプリケーションの2番目のインスタンスをボックスウィンドウにダウンロードして起動します。これにより、SPA の目的が部分的には損なわれます。 また、ダイアログボックスウィンドウで変数を変更するコードでは、同じ変数の作業ウィンドウバージョンは変更されません。 同様に、ダイアログボックスウィンドウには独自のセッションストレージがあります。これには、作業ウィンドウのコードからアクセスすることはできません。 ダイアログボックスと呼び出し先のホストページが、 `displayDialogAsync` サーバーに対して2つの異なるクライアントのように表示されています。 (ホストページについての通知については、「[ホストページからダイアログボックスを開く](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)」を参照してください)。

そのため、メソッドにルートを渡すと、実際には `displayDialogAsync` spa を持っていません。*同じ spa のインスタンスが2つ*あります。 さらに、作業ウィンドウインスタンス内のコードの多くは、そのインスタンスでは使用されず、ダイアログボックスインスタンス内のコードの多くは、そのインスタンスでは使用されません。 同じバンドルに 2 つの SPA があるようなものです。

#### <a name="microsoft-recommendations"></a>Microsoft の推奨事項

メソッドにクライアント側のルートを渡す代わりに `displayDialogAsync` 、次のいずれかの手順を実行することをお勧めします。

* ダイアログボックスで実行するコードが非常に複雑な場合は、2つの異なる SPAs 明示的に作成します。つまり、同じドメインの異なるフォルダーに2つの SPAs 持たせることができます。 ダイアログボックス内で1つの SPA が実行され、もう1つは、が呼び出されたダイアログボックスのホストページに `displayDialogAsync` あります。 
* ほとんどのシナリオでは、ダイアログボックスには単純なロジックのみが必要です。 このような場合、1つの HTML ページ (埋め込まれた、または参照された JavaScript を含む) を SPA のドメインにホストすることによって、プロジェクトが大幅に簡素化されます。 ページの URL を `displayDialogAsync` メソッドに渡します。 これは、単一ページアプリのリテラルの概念から deviating していることを意味します。Office ダイアログ API を使用している場合、実際には SPA の1つのインスタンスを持っていません。
