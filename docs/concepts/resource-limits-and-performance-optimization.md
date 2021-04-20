---
title: Office アドインのリソースの制限とパフォーマンスの最適化
description: CPU とメモリを含む、Office アドインプラットフォームのリソースの制限について説明します。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: 8c64e5a836d6b998ccd7022e71f595bb331bba8c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293332"
---
# <a name="resource-limits-and-performance-optimization-for-office-add-ins"></a>Office アドインのリソースの制限とパフォーマンスの最適化

ユーザーのベスト エクスペリエンスを実現するために、Office アドイン実行時の CPU コア、メモリの使用量、信頼性、および Outlook アドインの正規表現の評価の応答時間を一定以内に保つ必要があります。これらの実行時のリソース使用量の制限は、Windows と OS X 用の Office クライアントに適用され、モバイルアプリやブラウザーでは適用されません。

また、デスクトップやモバイル デバイス上のアドインについても、アドインの設計と実装でリソース使用量を最適化することによって、そのパフォーマンスを最適化できます。

## <a name="resource-usage-limits-for-add-ins"></a>アドインのリソース使用量の制限

実行時のリソース使用量の制限は、すべての種類の Office アドインに適用されます。これらの制限は、ユーザーのパフォーマンスを確保し、サービス拒否攻撃を軽減するのに役立ちます。 使用可能なデータの範囲を使用して、対象の Office アプリケーションで Office アドインをテストし、次の実行時の使用法の制限に対してパフォーマンスを測定します。

- **CPU コアの使用率**: 単一の CPU コアの使用率しきい値 90%、既定の 5 秒間隔で 3 回観測。

   Office クライアントが CPU コア使用率をチェックする既定の間隔は、5秒ごとです。 Office クライアントが、アドインの CPU コアの使用率をしきい値を超える場合は、ユーザーがアドインの実行を続行するかどうかを確認するメッセージを表示します。 ユーザーが続行することを選択した場合、Office クライアントは編集セッション中にユーザーに対して再度要求を行いません。 ユーザーが CPU を集中的に使用するアドインを実行する場合、この警告メッセージの表示を減らすには、管理者は **AlertInterval** レジストリ キーを使用する必要がある可能性があります。

- **メモリ使用量**: デバイスの利用可能な物理メモリに基づいて動的に決定される、既定のメモリ使用量しきい値。

   既定では、Office クライアントは、デバイス上の物理メモリの使用率が使用可能なメモリの80% を超えることを検出すると、アドインのメモリ使用量、コンテンツアドインと作業ウィンドウアドインのドキュメントレベル、および Outlook アドインのメールボックスレベルで、クライアントによって監視が開始されます。既定の5秒の間隔で、ドキュメントまたはメールボックスレベルのアドインの物理メモリ使用量が50% を超えた場合、クライアントはユーザーに警告を表示します。 このメモリ使用量の制限は、仮想メモリではなく物理を使用して、タブレットなどの RAM に制限があるデバイスのパフォーマンスを確保します。 管理者は、グローバル設定として **Memoryalertthreshold** Windows レジストリキーを使用して明示的な制限でこの動的設定を上書きできます。 ir は、グローバル設定として **alertinterval** キーを使用してアラート間隔を調整します。

- **クラッシュ許容度**: 既定の制限は、1 つのアドインにつき 4 回。

   管理者は、**RestartManagerRetryLimit** レジストリ キーを使用して、クラッシュのしきい値を調整できます。

- **アプリケーションのブロッキング**: アドインが応答しないままになる時間のしきい値は 5 秒間。

   これは、アドインと Office アプリケーションのユーザーエクスペリエンスに影響します。 この場合、Office アプリケーションは、ドキュメントまたはメールボックス (該当する場合) に対してすべてのアクティブなアドインを自動的に再起動し、アドインが応答しなくなったことをユーザーに警告します。 アドインが時間のかかるタスクを実行していて定期的に処理を発生させないときに、このしきい値に到達する場合があります。 ブロッキングが発生しないようにする手法があります。 管理者は、このしきい値を上書きすることはできません。

### <a name="outlook-add-ins"></a>Outlook アドイン

Outlook アドインが前述の CPU コア使用率、メモリ使用量、またはクラッシュ許容度のしきい値を超えると、そのアドインは Outlook で無効化されます。Exchange 管理センターにはそのアプリの無効状態が表示されます。

> [!NOTE]
> Outlook on the web やモバイル端末ではなく、Outlook リッチ クライアントによってのみ、リソース使用量をモニターする場合でも、リッチ クライアントが Outlook アドインを無効化すると、このアドインは Outlook on the web やモバイル端末でも無効化されます。

CPU コア、メモリ、および信頼性ルールだけでなく、Outlook アドインは次のアクティブ化のルールを監視する必要があります。

- **正規表現の応答時間**: Outlook で Outlook アドインのマニフェスト内のすべての正規表現を評価する時間の既定のしきい値は 1,000 ミリ秒。このしきい値を超えると、Outlook は後で評価を再試行します。

    管理者は、Windows レジストリでグループ ポリシーまたはアプリケーション固有の設定として **OutlookActivationAlertThreshold** 設定を使用して、この 1,000 ミリ秒の既定のしきい値を調節できます。

- **正規表現の再評価**: Outlook でマニフェスト内の正規表現を再評価する既定の制限は 3 回。 適用されるしきい値 (既定の 1,000 ミリ秒、または Windows レジストリに **OutlookActivationAlertThreshold** 設定が存在する場合はその設定で指定された値) を 3 回とも超えて評価に失敗すると、その Outlook アドインは Outlook で無効化されます。 Exchange 管理センターに [無効] の状態が表示され、Outlook リッチクライアントおよび Outlook on the web およびモバイルデバイスで使用するアドインが無効になっています。

    管理者は、Windows レジストリでグループ ポリシーまたはアプリケーション固有の設定として **OutlookActivationManagerRetryLimit** 設定を使用して、評価を再試行するこの回数を調節できます。

### <a name="excel-add-ins"></a>Excel アドイン

Excel アドインを作成している場合は、ブックを操作するときに以下のサイズ制限に注意してください。

- Excel on the web ではペイロードのサイズが要求と応答で 5 MB に制限されています。 その制限を超えると、`RichAPI.Error` がスローされます。
- 範囲は、取得操作に500万のセルに制限されます。

ユーザー入力がこれらの制限を超えていることが予想される場合は、必ずデータを確認してから、を呼び出してください `context.sync()` 。 必要に応じて、操作を小さな部分に分割します。 `context.sync()`各サブ操作を呼び出して、それらの操作が再度一括されないようにしてください。

これらの制限は、通常、大きな範囲を超えています。 アドインでは、範囲内のセルを戦略的に更新するために [Rangeareas](/javascript/api/excel/excel.rangeareas) を使用できる場合があります。 詳細については、「 [Excel アドインで複数の範囲を同時に操作](../excel/excel-add-ins-multiple-ranges.md) する」を参照してください。

### <a name="task-pane-and-content-add-ins"></a>作業ウィンドウ アドインとコンテンツ アドイン

いずれかのコンテンツまたは作業ウィンドウアドインが CPU コアまたはメモリ使用率のしきい値を超えている場合、またはクラッシュの許容範囲制限が超過している場合は、対応する Office アプリケーションによってユーザーに対して警告が表示されます。 この時点で、ユーザーは次のどちらかの処理を実行できます。

- アドインを再起動します。
- しきい値を超えたというそれ以降の警告をキャンセルします。理想的な対処としては、ユーザーはそのアドインをドキュメントから削除する必要があります。そのアドインの使用を続行すると、さらにパフォーマンスと安定性の問題が発生する可能性があります。  

## <a name="verifying-resource-usage-issues-in-the-telemetry-log"></a>テレメトリ ログでリソース使用量の問題を確認する

Office には、Office アドインでのリソースの使用に関する問題も含めて、ローカル コンピューター上で実行される Office ソリューションの一定のイベント (読み込む、開く、閉じる、およびエラー) の記録を保守するテレメトリ ログが用意されています。テレメトリ ログを設定してある場合は、Excel を使用して、ローカル ドライブ上の次の既定の場所にあるテレメトリ ログを開くことができます。

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
|8 |アドインのマニフェストがダウンロードされませんでした|重大|Office アプリケーションは、SharePoint カタログ、コーポレートカタログ、または AppSource から Office アドインのマニフェストファイルを読み込むことができませんでした。|
|9 |アドインのマークアップを解析できませんでした|重大|Office アプリケーションは Office アドインマニフェストを読み込みましたが、アプリの HTML マークアップを読み取れませんでした。|
|10 |アドインの CPU 使用率が高すぎます|重大|Office アドインは、限定された時間内に CPU リソースの 90% 超を使用しました。|
|15 |アドインは文字列検索のタイムアウトのため無効になっています||Outlook アドインは電子メールの件名とメッセージを検索して、それらを正規表現で表示するかどうかを決定します。**[File]** 列に記された Outlook アドインは、正規表現での一致を試みている最中に繰り返しタイムアウトしたため、Outlook によって無効にされました。|
|18 |アドインは正常に終了しました||Office アプリケーションは、Office アドインを正常に閉じることができました。|
|年|アドインで実行時エラーが発生しました|重大|Office アドインに、エラーの原因となる問題がありました。 詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して **Microsoft Office Alerts** ログを確認してください。|
|1280|アドインでライセンスを確認できませんでした|重大|Office アドインのライセンス情報を確認できないか、有効期限が切れている可能性があります。 詳細については、エラーが発生したコンピューター上で Windows イベント ビューアーを使用して **Microsoft Office Alerts** ログを確認してください。|

詳細については、「[テレメトリ ダッシュボードを展開する](/previous-versions/office/office-2013-resource-kit/jj219431(v=office.15))」および「[テレメトリ ログを使用した Office ファイルおよびカスタム ソリューションのトラブルシューティング](/office/client-developer/shared/troubleshooting-office-files-and-custom-solutions-with-the-telemetry-log)」を参照してください。

## <a name="design-and-implementation-techniques"></a>設計および実装上のテクニック

CPU 使用率、メモリ使用量、クラッシュ許容度、UI の応答性に対するリソース制限は、リッチ クライアント上で実行される Office アドインにのみ適用されますが、サポートするすべてのクライアントおよびデバイス上でアドインが十分なパフォーマンスを発揮するためには、これらのリソース使用量およびバッテリーの使用量を最適化することが重要になります。 アドインで長時間実行される処理があったり、大規模なデータ セットを処理したりする場合は、最適化が特に重要です。 次の一覧に、CPU 負荷の高い操作を小さなチャンクに分割して、アドインが過剰なリソース消費を避け、Office アプリケーションの応答性を維持できるようにする方法を示します。

- 制限のないデータセットからの大量のデータをアドインで読み取る必要があるシナリオでは、テーブルからデータを読み取る場合にページ付けを適用したり、またはより小さいサイズの読み取り操作に分割して 1 回の操作で処理するデータ量を小さくし、1 回の操作ですべてのデータを読み取ることがないようにします。 この操作を行うには、グローバルオブジェクトの [setTimeout](https://developer.mozilla.org/docs/Web/API/WindowOrWorkerGlobalScope/setTimeout) メソッドを使用して、入力と出力の時間を制限します。 It also handles the data in defined chunks instead of randomly unbounded data. 別の方法として、 [async](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/async_function) を使用して約束を処理する方法があります。

- アドインで CPU 使用率の高いアルゴリズムを使用して大量のデータを処理する場合は、Web Workers を使用してバックグラウンドで時間のかかるタスクを実行しつつ、フォアグラウンドで別のスクリプト (ユーザー インターフェイスへの進行状況の表示など) を実行できます。Web Workers は、ユーザー アクティビティをブロックせず、HTML ページの応答性を維持します。Web Workers の例については、「[ウェブ ワーカーの基本](https://www.html5rocks.com/tutorials/workers/basics/)」を参照してください。Web Workers API の詳細については、「[Web Workers](https://developer.mozilla.org/docs/Web/API/Web_Workers_API)」を参照してください。

- アドインで CPU 使用率の高いアルゴリズムを使用しているが、データの入出力を小さなセットに分割できる場合は、Web サービスの作成を検討します。データを Web サービスに渡して CPU の負荷をオフロードし、非同期コールバックを待機します。

- 想定する最大量のデータでアドインをテストして、アドインにおける処理をその最大量までに制限します。

### <a name="performance-improvements-with-the-application-specific-apis"></a>アプリケーション固有の Api によるパフォーマンスの向上

[アプリケーション固有の api モデルを使用](../develop/application-specific-api-model.md)する際のパフォーマンスのヒントは、Excel、OneNote、Visio、および Word 用のアプリケーション固有の api を使用する際のガイダンスを提供します。 要約すると、次のことを行う必要があります。

- [必要なプロパティのみを読み込む](../develop/application-specific-api-model.md#calling-load-without-parameters-not-recommended)。
- [同期 () の呼び出しの数を最小限に](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-sync-calls)します。 「」を参照してください。コードで呼び出しを管理する方法について詳しくは [、「ループでのコンテキストの同期」メソッドの使用を避けて](correlated-objects-pattern.md) `sync` ください。
- [作成するプロキシオブジェクトの数を最小限に](../develop/application-specific-api-model.md#performance-tip-minimize-the-number-of-proxy-objects-created)します。 次のセクションで説明するように、プロキシオブジェクトを不要することもできます。

#### <a name="untrack-unneeded-proxy-objects"></a>不要不要なプロキシオブジェクト

[プロキシオブジェクト](../develop/application-specific-api-model.md#proxy-objects) は、が呼び出されるまでメモリ内に保持 `RequestContext.sync()` されます。 大規模なバッチ操作では、アドインが 1 回のみ必要とするプロキシ オブジェクトが大量に生成されることがあります。それらのオブジェクトは、バッチの実行前にメモリから解放できます。

メソッドは、 `untrack()` オブジェクトをメモリから解放します。 このメソッドは、多くのアプリケーション固有の API プロキシオブジェクトに実装されています。 アドインがこのオブジェクトを使用して `untrack()` 実行された後に呼び出しを行うと、多数のプロキシオブジェクトを使用する場合のパフォーマンスが大幅に向上します。

> [!NOTE]
> `Range.untrack()` は、[ClientRequestContext.trackedObjects.remove(thisRange)](/javascript/api/office/officeextension.trackedobjects#remove-object-) のショートカットです。 プロキシ オブジェクトは、コンテキスト内の追跡対象オブジェクト リストから削除することで追跡解除できます。

次の Excel コードサンプルでは、一度に1つのセルを選択して、選択範囲にデータを入力します。 セルに値が追加されると、そのセルを表している範囲の追跡が解除されます。 10,000 から 20,000 個のセルの範囲を選択して、このコードを実行します。最初の実行では `cell.untrack()` の行を使用し、その後でこの行を削除して実行します。 `cell.untrack()` の行がないコードよりも、この行があるコードの方が高速になることがわかります。 また、クリーンアップの手順にかかる時間が短くなるため、その後の応答時間も速くなることがわかります。

```js
Excel.run(async (context) => {
    var largeRange = context.workbook.getSelectedRange();
    largeRange.load(["rowCount", "columnCount"]);
    await context.sync();

    for (var i = 0; i < largeRange.rowCount; i++) {
        for (var j = 0; j < largeRange.columnCount; j++) {
            var cell = largeRange.getCell(i, j);
            cell.values = [[i *j]];

            // Call untrack() to release the range from memory.
            cell.untrack();
        }
    }

    await context.sync();
});
```

不要オブジェクトに必要なものは、何千ものものを扱っている場合にのみ重要になることに注意してください。 ほとんどのアドインでは、プロキシオブジェクト追跡を管理する必要はありません。

## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Excel の JavaScript API を使用した、パフォーマンスの最適化](../excel/performance.md)
