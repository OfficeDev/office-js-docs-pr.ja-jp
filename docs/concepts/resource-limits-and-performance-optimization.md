---
title: Office アドインのリソースの制限とパフォーマンスの最適化
description: CPU やメモリなど、Office アドイン プラットフォームのリソース制限について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f9bec9579db1461f16d36d97646c4fce418c2e11
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889199"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Office アドインのリソースの制限とパフォーマンスの最適化

ユーザーのベスト エクスペリエンスを実現するために、Office アドイン実行時の CPU コア、メモリの使用量、信頼性、および Outlook アドインの正規表現の評価の応答時間を一定以内に保つ必要があります。これらの実行時のリソース使用量の制限は、Windows と OS X 用の Office クライアントに適用され、モバイルアプリやブラウザーでは適用されません。

また、デスクトップやモバイル デバイス上のアドインについても、アドインの設計と実装でリソース使用量を最適化することによって、そのパフォーマンスを最適化できます。

## <a name="resource-usage-limits-for-add-ins"></a>アドインのリソース使用量の制限

ランタイム リソースの使用制限は、すべての種類の Office アドインに適用されます。これらの制限は、ユーザーのパフォーマンスを確保し、サービス拒否攻撃を軽減するのに役立ちます。 可能な範囲のデータを使用して、対象の Office アプリケーションで Office アドインをテストし、次の実行時の使用制限に照らしてパフォーマンスを測定してください。

- **CPU コアの使用率**: 単一の CPU コアの使用率しきい値 90%、既定の 5 秒間隔で 3 回観測。

   Office クライアントが CPU コア使用率をチェックする既定の間隔は 5 秒ごとです。 Office クライアントがアドインの CPU コア使用率がしきい値を超えていると検出した場合は、ユーザーがアドインの実行を続行するかどうかを確認するメッセージが表示されます。 ユーザーが続行することを選択した場合、Office クライアントは、その編集セッション中にユーザーにもう一度要求しません。 ユーザーが CPU を集中的に使用するアドインを実行する場合、この警告メッセージの表示を減らすには、管理者は **AlertInterval** レジストリ キーを使用する必要がある可能性があります。

- **メモリ使用量**: デバイスの利用可能な物理メモリに基づいて動的に決定される、既定のメモリ使用量しきい値。

   既定では、Office クライアントがデバイス上の物理メモリ使用量が使用可能なメモリの 80% を超えていると検出すると、クライアントはアドインのメモリ使用量の監視を、コンテンツと作業ウィンドウ アドインのドキュメント レベル、Outlook アドインのメールボックス レベルで開始します。既定の間隔 5 秒で、ドキュメントまたはメールボックス レベルでアドインのセットの物理メモリ使用量が 50% を超えた場合、クライアントはユーザーに警告します。 このメモリ使用量の制限では、仮想メモリではなく物理メモリを使用して、タブレットなどの RAM が制限されたデバイスでのパフォーマンスを確保します。 管理者は **、MemoryAlertThreshold** Windows レジストリ キーをグローバル設定として使用し、グローバル設定として **AlertInterval** キーを使用してアラート間隔を調整することで、この動的設定を明示的な制限でオーバーライドできます。

- **クラッシュ許容度**: 既定の制限は、1 つのアドインにつき 4 回。

   管理者は、**RestartManagerRetryLimit** レジストリ キーを使用して、クラッシュのしきい値を調整できます。

- **アプリケーションのブロッキング**: アドインが応答しないままになる時間のしきい値は 5 秒間。

   これは、アドインと Office アプリケーションに対するユーザーのエクスペリエンスに影響します。 これが発生すると、Office アプリケーションはドキュメントまたはメールボックスのすべてのアクティブなアドイン (該当する場合) を自動的に再起動し、アドインが応答しなくなったユーザーに警告します。 アドインが時間のかかるタスクを実行していて定期的に処理を発生させないときに、このしきい値に到達する場合があります。 ブロッキングが発生しないようにする手法があります。 管理者は、このしきい値を上書きすることはできません。

### <a name="outlook-add-ins"></a>Outlook アドイン

Outlook アドインが前述の CPU コア使用率、メモリ使用量、またはクラッシュ許容度のしきい値を超えると、そのアドインは Outlook で無効化されます。Exchange 管理センターにはそのアプリの無効状態が表示されます。

> [!NOTE]
> Outlook on the web やモバイル端末ではなく、Outlook リッチ クライアントによってのみ、リソース使用量をモニターする場合でも、リッチ クライアントが Outlook アドインを無効化すると、このアドインは Outlook on the web やモバイル端末でも無効化されます。

CPU コア、メモリ、および信頼性ルールに加えて、Outlook アドインは、アクティブ化に関する次の規則を遵守する必要があります。

- **正規表現の応答時間**: Outlook で Outlook アドインのマニフェスト内のすべての正規表現を評価する時間の既定のしきい値は 1,000 ミリ秒。このしきい値を超えると、Outlook は後で評価を再試行します。

    管理者は、Windows レジストリでグループ ポリシーまたはアプリケーション固有の設定として **OutlookActivationAlertThreshold** 設定を使用して、この 1,000 ミリ秒の既定のしきい値を調節できます。

- **正規表現の再評価**: Outlook でマニフェスト内の正規表現を再評価する既定の制限は 3 回。 適用されるしきい値 (既定の 1,000 ミリ秒、または Windows レジストリに **OutlookActivationAlertThreshold** 設定が存在する場合はその設定で指定された値) を 3 回とも超えて評価に失敗すると、その Outlook アドインは Outlook で無効化されます。 Exchange 管理 センターには無効な状態が表示され、Outlook リッチ クライアントとOutlook on the webおよびモバイル デバイスで使用するためにアドインが無効になります。

    管理者は、Windows レジストリでグループ ポリシーまたはアプリケーション固有の設定として **OutlookActivationManagerRetryLimit** 設定を使用して、評価を再試行するこの回数を調節できます。

### <a name="excel-add-ins"></a>Excel アドイン

Excel アドインをビルドする場合は、ブックを操作するときに、次のサイズの制限に注意してください。

- Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。 その制限を超えると、`RichAPI.Error` がスローされます。
- 取得操作の範囲は 500 万セルに制限されています。

ユーザー入力がこれらの制限を超えると予想される場合は、呼び出す `context.sync()`前に必ずデータを確認してください。 必要に応じて、操作を小さな部分に分割します。 これらの操作が再度バッチ処理されないように、サブ操作ごとに必ず呼び出 `context.sync()` してください。

これらの制限は通常、大きな範囲で超えています。 アドインでは [、RangeAreas](/javascript/api/excel/excel.rangeareas) を使用して、より大きな範囲内のセルを戦略的に更新できる場合があります。 操作 `RangeAreas`の詳細については、「 [Excel アドインで複数の範囲を同時に操作する](../excel/excel-add-ins-multiple-ranges.md)」を参照してください。Excel でのペイロード サイズの最適化の詳細については、「 [ペイロード サイズの制限に関するベスト プラクティス](../excel/performance.md#payload-size-limit-best-practices)」を参照してください。

### <a name="task-pane-and-content-add-ins"></a>作業ウィンドウ アドインとコンテンツ アドイン

コンテンツまたは作業ウィンドウ アドインが、CPU コアまたはメモリ使用量の前のしきい値を超えた場合、またはクラッシュに対する許容範囲の制限を超えると、対応する Office アプリケーションにユーザーに対する警告が表示されます。 この時点で、ユーザーは次のどちらかの処理を実行できます。

- アドインを再起動します。
- しきい値を超えたというそれ以降の警告をキャンセルします。理想的な対処としては、ユーザーはそのアドインをドキュメントから削除する必要があります。そのアドインの使用を続行すると、さらにパフォーマンスと安定性の問題が発生する可能性があります。  

## <a name="verify-resource-usage-issues-in-the-telemetry-log"></a>テレメトリ ログでリソース使用状況の問題を確認する

Office には、Office アドインでのリソースの使用に関する問題も含めて、ローカル コンピューター上で実行される Office ソリューションの一定のイベント (読み込む、開く、閉じる、およびエラー) の記録を保守するテレメトリ ログが用意されています。 テレメトリ ログを設定している場合は、Excel を使用して、ローカル ドライブ上の次の既定の場所でテレメトリ ログを開くことができます。

`%Users%\<Current user>\AppData\Local\Microsoft\Office\15.0\Telemetry`

それぞれのアドインについてテレメトリ ログで追跡されるイベントごとに、そのイベントの発生日付/時刻、イベント ID、重大度、および短い説明的なタイトル、そのアドインのフレンドリ名と ID、イベントをログに記録したアプリケーションが記入されています。テレメトリ ログをリフレッシュすれば、現在の追跡済みイベントを確認できます。次の表は、テレメトリ ログで追跡された Outlook アドインの例を示しています。

|**日付/時刻**|**イベント ID**|**重大度**|**タイトル**|**ファイル**|**ID**|**アプリケーション**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|10/8/2012 5:57:10 PM|7 ||アドインのマニフェストが正常にダウンロードされました|重要人物|69cc567c-6737-4c49-88dd-123334943a22|Outlook|
|10/8/2012 5:57:01 PM|7 ||アドインのマニフェストが正常にダウンロードされました|LinkedIn|333bf46d-7dad-4f2b-8cf4-c19ddc78b723|Outlook|

次の表は、通常、Office アドインについてテレメトリ ログで追跡されるイベントを示しています。

|**イベント ID**|**タイトル**|**重大度**|**説明**|
|:-----|:-----|:-----|:-----|
|7 |アドインのマニフェストが正常にダウンロードされました||Office アドインのマニフェストが正常に読み込まれ、Office アプリケーションによって読み取られました。|
|8 |アドインのマニフェストがダウンロードされませんでした|重大|Office アプリケーションは、SharePoint カタログ、企業カタログ、または AppSource から Office アドインのマニフェスト ファイルを読み込めませんでした。|
|9 |アドインのマークアップを解析できませんでした|重大|Office アプリケーションは Office アドイン マニフェストを読み込みましたが、アプリの HTML マークアップを読み取ることができませんでした。|
|10|アドインの CPU 使用率が高すぎます|重大|Office アドインは、限定された時間内に CPU リソースの 90% 超を使用しました。|
|15|アドインは文字列検索のタイムアウトのため無効になっています||Outlook アドインは電子メールの件名とメッセージを検索して、それらを正規表現で表示するかどうかを決定します。**[File]** 列に記された Outlook アドインは、正規表現での一致を試みている最中に繰り返しタイムアウトしたため、Outlook によって無効にされました。|
|18 |アドインは正常に終了しました||Office アプリケーションは、Office アドインを正常に閉じることができました。|
|19|アドインで実行時エラーが発生しました|重大|Office アドインに、エラーの原因となる問題がありました。 詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して **Microsoft Office Alerts** ログを確認してください。|
|20|アドインでライセンスを確認できませんでした|重大|Office アドインのライセンス情報を確認できないか、有効期限が切れている可能性があります。 詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して **Microsoft Office Alerts** ログを確認してください。|

詳細については、「[テレメトリ ダッシュボードを展開する](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15))」および「[テレメトリ ログを使用した Office ファイルおよびカスタム ソリューションのトラブルシューティング](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)」を参照してください。

## <a name="design-and-implementation-techniques"></a>設計および実装上のテクニック

CPU 使用率、メモリ使用量、クラッシュ許容度、UI の応答性に対するリソース制限は、リッチ クライアント上で実行される Office アドインにのみ適用されますが、サポートするすべてのクライアントおよびデバイス上でアドインが十分なパフォーマンスを発揮するためには、これらのリソース使用量およびバッテリーの使用量を最適化することが重要になります。 アドインで長時間実行される処理があったり、大規模なデータ セットを処理したりする場合は、最適化が特に重要です。 次の一覧では、アドインが過度なリソース消費を回避し、Office アプリケーションが応答性を維持できるように、CPU 集中型またはデータ集中型の操作をより小さなチャンクに分割するいくつかの手法を示します。

- 制限のないデータセットからの大量のデータをアドインで読み取る必要があるシナリオでは、テーブルからデータを読み取る場合にページ付けを適用したり、またはより小さいサイズの読み取り操作に分割して 1 回の操作で処理するデータ量を小さくし、1 回の操作ですべてのデータを読み取ることがないようにします。 これを行うには、グローバル オブジェクトの [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) メソッドを使用して、入力と出力の期間を制限します。 It also handles the data in defined chunks instead of randomly unbounded data. もう 1 つのオプションは、 [async](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) を使用して Promise を処理することです。

- アドインで CPU 使用率の高いアルゴリズムを使用して大量のデータを処理する場合は、Web Workers を使用してバックグラウンドで時間のかかるタスクを実行しつつ、フォアグラウンドで別のスクリプト (ユーザー インターフェイスへの進行状況の表示など) を実行できます。Web Workers は、ユーザー アクティビティをブロックせず、HTML ページの応答性を維持します。Web Workers の例については、「[ウェブ ワーカーの基本](https://www.html5rocks.com/tutorials/workers/basics/)」を参照してください。Web Workers API の詳細については、「[Web Workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API)」を参照してください。

- アドインで CPU 使用率の高いアルゴリズムを使用しているが、データの入出力を小さなセットに分割できる場合は、Web サービスの作成を検討します。データを Web サービスに渡して CPU の負荷をオフロードし、非同期コールバックを待機します。

- 想定する最大量のデータでアドインをテストして、アドインにおける処理をその最大量までに制限します。

### <a name="performance-improvements-with-the-application-specific-apis"></a>アプリケーション固有の API によるパフォーマンスの向上

[アプリケーション固有の API モデルの使用に関するパフォーマンスに](../develop/application-specific-api-model.md)関するヒントでは、Excel、OneNote、Visio、および Word 用のアプリケーション固有の API を使用する場合のガイダンスを提供します。 要約すると、次のことを行う必要があります。

- [必要なプロパティのみを読み込みます](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended)。
- [sync() 呼び出しの数を最小限に抑えます](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls)。 コード内の呼び出しを管理`sync`する方法の詳細については、[ループで context.sync メソッドを使用しないでください](correlated-objects-pattern.md)。
- [作成されたプロキシ オブジェクトの数を最小限に抑えます](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created)。 次のセクションで説明するように、プロキシ オブジェクトの追跡を解除することもできます。

#### <a name="untrack-unneeded-proxy-objects"></a>不要なプロキシ オブジェクトを追跡解除する

[プロキシ オブジェクト](../develop/application-specific-api-model.md#proxy-objects) は、呼び出されるまで `RequestContext.sync()` メモリ内に保持されます。 大規模なバッチ操作では、アドインが 1 回のみ必要とするプロキシ オブジェクトが大量に生成されることがあります。それらのオブジェクトは、バッチの実行前にメモリから解放できます。

このメソッドは `untrack()` 、オブジェクトをメモリから解放します。 このメソッドは、多くのアプリケーション固有の API プロキシ オブジェクトに実装されます。 アドインがオブジェクトで実行された後に呼び出すと `untrack()` 、多数のプロキシ オブジェクトを使用する場合に、パフォーマンス上の利点が顕著に得られます。

> [!NOTE]
> `Range.untrack()` は、[ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#office-officeextension-trackedobjects-remove-member(1)) のショートカットです。 プロキシ オブジェクトは、コンテキスト内の追跡対象オブジェクト リストから削除することで追跡解除できます。

次の Excel コード サンプルは、選択した範囲にデータを一度に 1 セルずつ入力します。 セルに値が追加されると、そのセルを表している範囲の追跡が解除されます。 10,000 から 20,000 個のセルの範囲を選択して、このコードを実行します。最初の実行では `cell.untrack()` の行を使用し、その後でこの行を削除して実行します。 `cell.untrack()` の行がないコードよりも、この行があるコードの方が高速になることがわかります。 また、クリーンアップの手順にかかる時間が短くなるため、その後の応答時間も速くなることがわかります。

```js
Excel.run(async (context) => {
    const largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (let i = 0; i < largeRange.rowCount; i++) {
        for (let j = 0; j < largeRange.columnCount; j++) {
            let cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

オブジェクトを追跡解除する必要は、何千ものオブジェクトを処理する場合にのみ重要であることに注意してください。 ほとんどのアドインでは、プロキシ オブジェクトの追跡を管理する必要はありません。

## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Excel の JavaScript API を使用した、パフォーマンスの最適化](../excel/performance.md)
