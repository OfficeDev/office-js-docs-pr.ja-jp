---
title: Office アドインのランタイム
description: Office アドインで使用されるランタイムについて説明します。
ms.date: 08/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8d28f6db028d2f4c7036db51ccc5dbcc2144bdf3
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616043"
---
# <a name="runtimes-in-office-add-ins"></a>Office アドインのランタイム

Office アドインは、Office に埋め込まれたランタイムで実行されます。 解釈される言語として、JavaScript は JavaScript エンジンで実行する必要があります。 シングルスレッドの同期言語である JavaScript には、同時実行のための固有の容量がありません。ただし、最新の JavaScript エンジンでは、ホスト オペレーティング システムから同時操作 (ネットワーク通信を含む) を要求し、応答として OS からデータを受信できます。 この種のエンジンにより、JavaScript は *効果的に* 非同期になります。 この記事では、この種のエンジンをランタイムと呼 *びます*。 [Node.js](https://nodejs.org) ブラウザーと最新のブラウザーは、このようなランタイムの例です。 

## <a name="types-of-runtimes"></a>ランタイムの種類

Office アドインで使用されるランタイムには、次の 2 種類があります。

- **JavaScript 専用ランタイム**: [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、 [Full CORS (クロスオリジン リソース共有)](https://developer.mozilla.org/docs/Web/HTTP/CORS)、およびクライアント側のデータストレージのサポートを補完する JavaScript エンジン。 ( [ローカル ストレージ](https://developer.mozilla.org/docs/Web/API/Window/localStorage) や Cookie はサポートされていません)。) 
- **ブラウザー ランタイム**: JavaScript 専用ランタイムのすべての機能が含まれており、 [ローカル ストレージ](https://developer.mozilla.org/docs/Web/API/Window/localStorage)、HTML を [レンダリングするレンダリング エンジン](https://developer.mozilla.org/docs/Glossary/Rendering_engine) 、Cookie のサポートが追加されます。

これらの型の詳細については、この記事の後半で [JavaScript 専用ランタイム](#javascript-only-runtime) と [ブラウザー ランタイムを参照してください](#browser-runtime)。

次の表は、アドインで使用されるランタイムの種類ごとに考えられる機能を示しています。 

> [!NOTE]
> 使用するランタイムの種類の選択は、Microsoft がいつでも変更できる実装の詳細です。 Office JavaScript ライブラリでは、特定の機能に対して常に同じ種類のランタイムが使用されるとは限りません。また、アドイン アーキテクチャでもこれを想定しないでください。

| ランタイムの種類 | アドイン機能 |
|:-----|:-----|
| JavaScript のみ | Excel [カスタム関数](../excel/custom-functions-overview.md)</br>(ランタイムが[共有](#shared-runtime)されている場合、またはアドインがOffice on the webで実行されている場合を除く)</br></br>[Outlook イベント ベースのタスク](../outlook/autolaunch.md)</br>(アドインが Outlook on Windows で実行されている場合のみ)|
| ブラウザー | [作業ウィンドウ](../design/task-pane-add-ins.md)</br></br>[ダイアログ](../develop/dialog-api-in-office-add-ins.md)</br></br>[function コマンド](../design/add-in-commands.md#types-of-add-in-commands)</br></br>Excel [カスタム関数](../excel/custom-functions-overview.md)</br>(ランタイムが[共有](#shared-runtime)されている場合、またはアドインがOffice on the webで実行されている場合)</br></br>[Outlook イベント ベースのタスク](../outlook/autolaunch.md)</br>(アドインが Outlook on Mac または Outlook on the webで実行されている場合)|

次の表は、アドインのさまざまな可能な機能に使用されるランタイムの種類によって編成された同じ情報を示しています。

| アドイン機能 | Windows でのランタイムの種類 | Mac でのランタイムの種類 | Web 上のランタイムの種類 |
|:-----|:-----|:-----|:-----|
|Excel のカスタム関数 | JavaScript のみ</br>(ただし、ランタイムが共有されている場合は *ブラウザー* )|JavaScript のみ</br>(ただし、ランタイムが共有されている場合は *ブラウザー* )| ブラウザー |
|Outlook イベント ベースのタスク | JavaScript のみ | ブラウザー | ブラウザー |
|作業ウィンドウ | ブラウザー | ブラウザー | ブラウザー |
|ダイアログ | ブラウザー | ブラウザー | ブラウザー |
|function コマンド | ブラウザー | ブラウザー | ブラウザー |


Office on the webでは、すべてが常にブラウザー型ランタイムで実行されます。 実際、1 つの例外を除き、Web 上のアドイン内のすべての処理は *、同じ* ブラウザー プロセス (ユーザーがOffice on the webを開いたブラウザー プロセス) で実行されます。 例外は、 [Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) の呼び出しでダイアログが開き [、DialogOptions.displayInIFrame](/javascript/api/office/office.dialogoptions#office-office-dialogoptions-displayiniframe-member) オプションが渡 *されず* に `true`設定されている場合です。 オプションが渡されない (既定値 `false` を持つ) 場合、ダイアログは独自のプロセスで開きます。 [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) メソッドと [OfficeRuntime.DisplayWebDialogOptions.displayInIFrame](/javascript/api/office-runtime/officeruntime.displaywebdialogoptions#office-runtime-officeruntime-displaywebdialogoptions-displayiniframe-member) オプションにも同じ原則が適用されます。

アドインが Web 以外のプラットフォームで実行されている場合は、次の原則が適用されます。

- ダイアログは、独自のランタイム プロセスで実行されます。 
- Outlook イベント ベースのタスクは、独自のランタイム プロセスで実行されます。 
- 既定では、作業ウィンドウ、関数コマンド、Excel カスタム関数は、それぞれ独自のランタイム プロセスで実行されます。 ただし、一部の Office ホスト アプリケーションでは、アドイン マニフェストを構成して、2 つまたは 3 つすべてを同じランタイムで実行できます。 [「共有ランタイム](#shared-runtime)」を参照してください。

ホスト Office アプリケーションとアドインで使用される機能によっては、アドインに多くのランタイムが存在する可能性があります。 通常、それぞれが独自のプロセスで実行されますが、必ずしも同時に実行されるとは限りません。 次に例を示します。

- ランタイムを共有せず、次の機能を含む PowerPoint または Word アドインには、3 つのランタイムがあります。

  - 作業ウィンドウ
  - 関数コマンド
  - ダイアログ (作業ウィンドウまたは関数コマンドからダイアログを起動できます)。 
  
      > [!NOTE]
      > 複数のダイアログを同時に開くのは良い方法ではありませんが、アドインで作業ウィンドウから 1 つを開き、関数コマンドから別のダイアログを同時に開ける場合、このアドインには 4 つのランタイムがあります。 作業ウィンドウと関数コマンドの特定の呼び出しでは、一度に開いているダイアログを 1 つだけ使用できます。ただし、関数コマンドが複数回呼び出された場合は、呼び出しごとに先行タスクの上に新しいダイアログが開き、多くのランタイムが発生する可能性があります。 この一覧の残りの部分では、複数の開いているダイアログの可能性は無視されます。

- ランタイムを共有せず、次の機能を含む Excel アドインには、 *4 つの* ランタイムがあります。

  - 作業ウィンドウ
  - 関数コマンド
  - カスタム関数
  - ダイアログ (作業ウィンドウ、関数コマンド、またはカスタム関数からダイアログを起動できます)。

- 同じ機能を備え、作業ウィンドウ、関数コマンド、およびカスタム関数全体で同じランタイムを共有するように構成された Excel アドインには、 *2 つの* ランタイムがあります。 共有ランタイムでは、一度に開くことができるダイアログは 1 つだけです。
- ダイアログがなく、作業ウィンドウ、関数コマンド、およびカスタム関数で同じランタイムを共有するように構成されている点を除き、同じ機能を備えた Excel アドインには、 *1 つの* ランタイムがあります。
- 次の機能を備えた Outlook アドインには、 *4 つの* ランタイムがあります。 (Outlook ではランタイムを共有できません。)

  - 作業ウィンドウ
  - 関数コマンド
  - イベント ベースのタスク
  - ダイアログ (ダイアログは作業ウィンドウまたは関数コマンドから起動できますが、イベント ベースのタスクからは起動できません)。

## <a name="share-data-across-runtimes"></a>ランタイム間でデータを共有する

> [!NOTE]
> - アドインがOffice on the webでのみ使用され、オプションが設定`true`されたダイアログ`displayInIFrame`が開かないことがわかっている場合は、このセクションを無視できます。 アドイン内のすべてが同じランタイム プロセスで実行されるため、グローバル変数を使用して機能間でデータを共有できます。
> - 前述のように、 [ランタイムの種類](#types-of-runtimes)で説明したように、機能で使用されるランタイムの種類は、プラットフォームによって部分的に異なります。 プラットフォームに基づいて分岐するアドイン コードを使用しないようにすることをお勧めします。そのため、このセクションのガイダンスでは、クロスプラットフォームで動作する手法をお勧めします。 分岐コードが必要なケースは、次に示す 1 つだけです。 

Excel、PowerPoint、および Word アドインの場合、ダイアログを除く 2 つ以上の機能でデータを共有する必要がある場合は、 [共有ランタイム](#shared-runtime) を使用します。 Outlook またはランタイムを共有できないシナリオでは、別の方法が必要です。 個別のランタイム プロセスにあるアドインの部分は、グローバル データを自動的に共有せず、アドインの Web アプリケーション サーバーによって個別のセッションとして扱われるため、 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) を使用してそれらの間でデータを共有することはできません。 *次のガイダンスでは、共有ランタイムを使用していないことを前提としています。*

- [Office.ui.messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) メソッドと [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) メソッドを使用して、ダイアログとその親作業ウィンドウ、関数コマンド、またはカスタム関数の間でデータを渡します。 

    > [!NOTE]
    > `OfficeRuntime.storage`ダイアログではメソッドを呼び出すことができないので、これはダイアログと別のランタイム間でデータを共有するためのオプションではありません。 

- 作業ウィンドウと関数コマンドの間でデータを共有するには、同じ特定の[配信元](https://developer.mozilla.org/docs/Glossary/Origin)にアクセスするすべてのランタイムで共有される [Window.localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) にデータを格納します。 
    > [!NOTE]
    > LocalStorage は JavaScript 専用ランタイムではアクセスできないため、Excel カスタム関数では使用できません。 また、Outlook イベント ベースのタスクとデータを共有するためにも使用できません (これらのタスクは一部のプラットフォームで JavaScript 専用ランタイムを使用するため)。

    > [!TIP]
    > データインは `Window.localStorage` アドインのセッション間で保持され、同じ配信元を持つアドインによって共有されます。 これらの特性はどちらも、多くの場合、アドインでは望ましくありません。 
    >
    > - 特定のアドインの各セッションが新しく開始されるようにするには、アドインの起動時に [Window.localStorage.clear](https://developer.mozilla.org/docs/Web/API/Storage/clear) メソッドを呼び出します。 
    > - 保存されている一部の値を保持できますが、他の値を再初期化するには、初期値にリセットする必要がある項目ごとにアドインが開始されるときに [Window.localStorage.setItem](https://developer.mozilla.org/docs/Web/API/Storage/setItem) を使用します。 
    > - アイテム全体を削除するには、 [Window.localStorage.removeItem](https://developer.mozilla.org/docs/Web/API/Storage/removeItem) を呼び出します。

- Excel カスタム関数とその他のランタイム間でデータを共有するには、 [OfficeRuntime.storage を使用します](/javascript/api/office-runtime/officeruntime.storage)。
- Outlook イベント ベースのタスクと作業ウィンドウまたは関数コマンドの間でデータを共有するには、 [Office.context.platform](/javascript/api/office/office.context#office-office-context-platform-member) プロパティの値でコードを分岐する必要があります。 

    - 値が (Windows) の場合は、`PC`[Office.sessionData](/javascript/api/outlook/office.sessiondata) API を使用してデータを格納および取得します。
    - 値が指定されている場合は、 `Mac`この一覧で前述したように使用 `Window.localStorage` します。

データを共有するその他の方法には、次のものがあります。

- すべてのランタイムからアクセスできるオンライン データベースに共有データを格納します。
- アドインのドメインの Cookie に共有データを格納して、ブラウザー ランタイム間で共有します。 JavaScript 専用ランタイムでは、Cookie はサポートされていません。

詳細については、「[アドインの状態と設定を保持する」と「Outlook アドインの](../develop/persisting-add-in-state-and-settings.md)[状態と設定を管理する](../outlook/manage-state-and-settings-outlook.md)」を参照してください。

## <a name="javascript-only-runtime"></a>JavaScript 専用ランタイム

Office アドインで使用される JavaScript 専用ランタイムは、[React Native用に](https://reactnative.dev/)最初に作成されたオープンソース ランタイムの変更です。 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、[Full CORS (クロスオリジン リソース共有)](https://developer.mozilla.org/docs/Web/HTTP/CORS)、[OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) のサポートを補完する JavaScript エンジンが含まれています。 レンダリング エンジンがなく、Cookie や [ローカル ストレージ](https://developer.mozilla.org/docs/Web/API/Window/localStorage)はサポートされていません。

この種類のランタイムは、カスタム関数が [ランタイムを共有](#shared-runtime)している場合を *除き*、Office on Windows の Outlook イベント ベースのタスクと Excel カスタム関数でのみ使用されます。 

- Excel カスタム関数に使用すると、ワークシートが再計算されるか、カスタム関数が計算されたときにランタイムが起動します。 ブックが閉じられるまでシャットダウンされません。  
- Outlook イベント ベースのタスクで使用すると、イベントが発生したときにランタイムが起動します。 次の 1 つ目が発生すると終了します。

  - イベント ハンドラーは、そのイベント パラメーターの `completed` メソッドを呼び出します。
  - トリガー イベントから 5 分が経過しました。
  - ユーザーは、イベントがトリガーされたウィンドウ (メッセージ作成ウィンドウなど) からフォーカスを変更します。

JavaScript ランタイムでは、ブラウザー ランタイムよりも少ないメモリを使用して起動しますが、機能は少なくなります。

## <a name="browser-runtime"></a>ブラウザー ランタイム

Office アドインは、Office が実行されているプラットフォーム (Web、Mac、または Windows) と、Windows と Office のバージョンとビルドに応じて、異なるブラウザーの種類のランタイムを使用します。 たとえば、ユーザーが FireFox ブラウザーでOffice on the webを実行している場合、Firefox ランタイムが使用されます。 ユーザーが Office on Mac を実行している場合は、Safari ランタイムが使用されます。 ユーザーが Windows で Office を実行している場合は、Windows と Office のバージョンに応じて、Edge または Internet Explorer がランタイムを提供します。 詳細については、 [Office アドインで使用されるブラウザーを参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

これらのランタイムはすべて HTML レンダリング エンジンを含み、 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)、 [フル CORS (クロスオリジン リソース共有)](https://developer.mozilla.org/docs/Web/HTTP/CORS)、 [ローカル ストレージ](https://developer.mozilla.org/docs/Web/API/Window/localStorage)、Cookie のサポートを提供します。 

ブラウザーランタイムの有効期間は、実装する機能と、それが共有されているかどうかによって異なります。

- 作業ウィンドウを含むアドインが起動されると、既に実行されている共有ランタイムでない限り、ブラウザー ランタイムが起動します。 共有ランタイムの場合は、ドキュメントが閉じられるとシャットダウンされます。 共有ランタイムでない場合は、作業ウィンドウが閉じられるとシャットダウンされます。
- ダイアログが開かれると、ブラウザー ランタイムが起動します。 ダイアログが閉じられると、シャットダウンされます。
- 関数コマンドが実行されると (ユーザーがボタンまたはメニュー項目を選択したときに発生します)、ブラウザー ランタイムは、既に実行されている共有ランタイムでない限り開始されます。 共有ランタイムの場合は、ドキュメントが閉じられるとシャットダウンされます。 共有ランタイムでない場合は、次の最初の実行時にシャットダウンされます。
 
  - 関数コマンドは、そのイベント パラメーターの `completed` メソッドを呼び出します。
  - トリガー イベントから 5 分が経過しました。 (関数コマンドでダイアログが開かれた場合でも、親ランタイムがタイムアウトしても開かれている場合は、ダイアログが閉じられるまでダイアログ ランタイムは実行されたままになります)。

- Excel カスタム関数が共有ランタイムを使用している場合、他の理由で共有ランタイムがまだ開始されていない場合、カスタム関数が計算されたときにブラウザー型ランタイムが開始されます。 ドキュメントが閉じられると、シャットダウンされます。

> [!NOTE]
> ランタイムが [共有](#shared-runtime)されている場合、アドインをシャットダウンせずに、コードで作業ウィンドウを閉じることが可能です。 詳細については、「 [Office アドインの作業ウィンドウを表示または非表示にする](../develop/show-hide-add-in.md) 」を参照してください。

ブラウザー ランタイムには、JavaScript 専用ランタイムよりも多くの機能がありますが、起動速度が低下し、より多くのメモリが使用されます。

### <a name="shared-runtime"></a>共有ランタイム

"共有ランタイム" は、ランタイムの種類ではありません。 これは、アドインの機能によって共有されている [ブラウザー型のランタイム](#browser-runtime) を指します。それ以外の場合は、それぞれが独自のランタイムを持ちます。 具体的には、ランタイムを共有するためにアドインの作業ウィンドウと関数コマンドを構成するオプションがあります。 Excel アドインでは、作業ウィンドウまたは関数コマンドまたはその両方のランタイムを共有するようにカスタム関数を構成することもできます。 これを行うと、カスタム関数は [JavaScript 専用](#javascript-only-runtime) ランタイムではなくブラウザー型ランタイムで実行されます。それ以外の場合と同様です。 [共有ランタイムを使用するようにアドインを構成する方法と、共有ランタイムを使用するように](../develop/configure-your-add-in-to-use-a-shared-runtime.md)アドインを構成する手順については、「共有ランタイムを使用するようにアドインを構成する」を参照してください。 簡単に言うと、JavaScript のみのランタイムでは、メモリの使用量が少なく、起動速度は速くなりますが、機能は少なくなります。

> [!NOTE]
> - ランタイムは、Excel、PowerPoint、Word でのみ共有できます。 
> - ランタイムを共有するようにダイアログを構成することはできません。 各ダイアログには常に独自のダイアログがあります。ただし、ダイアログがOffice on the webで起動され、オプションが `displayInIFrame` [ `true`.
> - 共有ランタイムでは、元の Microsoft Edge WebView (EdgeHTML) ランタイムは使用されません。 WebView2 で Microsoft Edge を使用するための条件 (Chromium ベース) が満たされている場合 ([Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)で指定されている場合)、そのランタイムが使用されます。 それ以外の場合は、Internet Explorer 11 ランタイムが使用されます。